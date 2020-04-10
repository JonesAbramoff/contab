VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MargContrOcx 
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11505
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   6600
      Index           =   3
      Left            =   180
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   11190
      Begin VB.TextBox AnaliseLinDescricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   60
         MaxLength       =   255
         TabIndex        =   45
         Top             =   630
         Width           =   2280
      End
      Begin VB.CommandButton BotaoRecalcularColuna 
         Caption         =   "Recalcular Coluna Corrente"
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
         Left            =   592
         TabIndex        =   33
         Top             =   6195
         Width           =   2850
      End
      Begin VB.CommandButton BotaoRecalcularPlanilha 
         Caption         =   "Limpar e Recalcular Toda a Planilha"
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
         Left            =   7357
         TabIndex        =   9
         Top             =   6180
         Width           =   3555
      End
      Begin VB.CommandButton BotaoLimpaRecalculaColuna 
         Caption         =   "Limpar e Recalcular Coluna Corrente"
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
         Left            =   3622
         TabIndex        =   8
         Top             =   6195
         Width           =   3555
      End
      Begin MSMask.MaskEdBox Valor6 
         Height          =   255
         Left            =   7155
         TabIndex        =   37
         Top             =   645
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor5 
         Height          =   255
         Left            =   6285
         TabIndex        =   38
         Top             =   630
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor8 
         Height          =   255
         Left            =   9030
         TabIndex        =   39
         Top             =   660
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor7 
         Height          =   255
         Left            =   8085
         TabIndex        =   40
         Top             =   660
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor3 
         Height          =   255
         Left            =   4470
         TabIndex        =   41
         Top             =   645
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor4 
         Height          =   255
         Left            =   5385
         TabIndex        =   42
         Top             =   630
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor2 
         Height          =   255
         Left            =   3585
         TabIndex        =   43
         Top             =   645
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor1 
         Height          =   255
         Left            =   2670
         TabIndex        =   44
         Top             =   615
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridAnalise 
         Height          =   6120
         Left            =   30
         TabIndex        =   46
         Top             =   15
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   10795
         _Version        =   393216
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   6600
      Index           =   1
      Left            =   195
      TabIndex        =   14
      Top             =   765
      Width           =   11160
      Begin VB.Frame FrameFaturamento 
         Caption         =   "Faturamento"
         Height          =   1665
         Left            =   1590
         TabIndex        =   23
         Top             =   3585
         Width           =   8340
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1800
            TabIndex        =   34
            Top             =   900
            Width           =   2130
         End
         Begin VB.ComboBox FilialFaturamento 
            Height          =   315
            ItemData        =   "MargContrOcx.ctx":0000
            Left            =   1800
            List            =   "MargContrOcx.ctx":0002
            TabIndex        =   5
            Top             =   405
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   300
            Left            =   5175
            TabIndex        =   6
            Top             =   412
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Index           =   2
            Left            =   495
            TabIndex        =   35
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label LabelFilialFat 
            AutoSize        =   -1  'True
            Caption         =   "Filial Faturamento:"
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
            Left            =   135
            TabIndex        =   31
            Top             =   465
            Width           =   1575
         End
         Begin VB.Label LabelVendedor 
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
            Left            =   4275
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   465
            Width           =   885
         End
      End
      Begin VB.Frame FrameProduto 
         Caption         =   "Produto"
         Height          =   1440
         Left            =   1575
         TabIndex        =   18
         Top             =   1800
         Width           =   8340
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   1845
            TabIndex        =   3
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1845
            TabIndex        =   4
            Top             =   930
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelUM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3375
            TabIndex        =   22
            Top             =   930
            Width           =   720
         End
         Begin VB.Label LabelQuantidade 
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
            Height          =   195
            Left            =   675
            TabIndex        =   21
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label LabelProduto 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   990
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3375
            TabIndex        =   19
            Top             =   360
            Width           =   4245
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   960
         Left            =   1575
         TabIndex        =   15
         Top             =   600
         Width           =   8340
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4980
            TabIndex        =   2
            Top             =   375
            Width           =   2190
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   1845
            TabIndex        =   1
            Top             =   375
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelFilial 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4440
            TabIndex        =   17
            Top             =   435
            Width           =   480
         End
         Begin VB.Label LabelCliente 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1095
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   16
            Top             =   435
            Width           =   660
         End
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   6600
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   750
      Visible         =   0   'False
      Width           =   11145
      Begin VB.Frame FrameDVV 
         Caption         =   "DVV"
         Height          =   4905
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   8625
         Begin VB.TextBox DVVValor3 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6345
            MaxLength       =   255
            TabIndex        =   29
            Top             =   2250
            Width           =   1200
         End
         Begin VB.TextBox DVVValor2 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5085
            MaxLength       =   255
            TabIndex        =   28
            Top             =   2250
            Width           =   1200
         End
         Begin VB.TextBox DVVValor1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3825
            MaxLength       =   255
            TabIndex        =   27
            Top             =   2250
            Width           =   1200
         End
         Begin VB.TextBox DVVDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   315
            MaxLength       =   255
            TabIndex        =   26
            Top             =   2250
            Width           =   3495
         End
         Begin MSFlexGridLib.MSFlexGrid GridDVV 
            Height          =   4335
            Left            =   165
            TabIndex        =   7
            Top             =   315
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   7646
            _Version        =   393216
         End
      End
   End
   Begin VB.CommandButton BotaoInsumos 
      Caption         =   "Insumos..."
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
      Left            =   7665
      TabIndex        =   36
      Top             =   180
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8955
      ScaleHeight     =   450
      ScaleWidth      =   1710
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   75
      Width           =   1770
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1170
         Picture         =   "MargContrOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   630
         Picture         =   "MargContrOcx.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   90
         Picture         =   "MargContrOcx.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabOpcao 
      Height          =   7080
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   12488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Despesas Variáveis de Venda"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Análise"
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
Attribute VB_Name = "MargContrOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'grid dvv
Dim iGrid_DVVDescricao_Col As Integer
Dim iGrid_DVVValor1_Col As Integer
Dim iGrid_DVVValor2_Col As Integer
Dim iGrid_DVVValor3_Col As Integer

'grid analise
Dim iGrid_Analise_Col As Integer
Dim iGrid_Valor1_Col As Integer
Dim iGrid_Valor2_Col As Integer
Dim iGrid_Valor3_Col As Integer
Dim iGrid_Valor4_Col As Integer
Dim iGrid_Valor5_Col As Integer
Dim iGrid_Valor6_Col As Integer
Dim iGrid_Valor7_Col As Integer
Dim iGrid_Valor8_Col As Integer

'controle de frame
Dim giFrameAtual As Integer

'controle de alteracao
Dim bAlterouTabPrincipal As Boolean
Dim iQtdeAlterada As Integer
Dim iProdutoAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFilialCliAlterada As Integer
Dim iFilialFatAlterada As Integer
Dim iVendedorAlterado As Integer
Dim iAlterado As Integer

'controle das tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DVV = 2
Private Const TAB_Analise = 3

'obj global com os dados da tela
Dim gobjMargContr As ClassMargContr

'obj dos grids
Dim objGridDVV As AdmGrid
Dim objGridAnalise As AdmGrid

'eventos do browser
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim bGridAnaliseCarregado As Boolean
Dim bGridDVVCarregado As Boolean

Private gobjTelaComissoes As Object
Private gcolComissoes As Collection
Private gdPrecoComissoes As Double

Public Function Trata_Parametros() As Long
'não espera parametros
    Trata_Parametros = SUCESSO
End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    bGridAnaliseCarregado = False
    bGridDVVCarregado = False
    
    'instancia o objMargContr
    Set gobjMargContr = New ClassMargContr
    
    'instancia os objs do grid
    Set objGridDVV = New AdmGrid
    Set objGridAnalise = New AdmGrid
    
    'instancia os obj dos browsers
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    'coloca o 1º tab como atual
    giFrameAtual = TAB_Identificacao
    
    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 119515
    
    'carrega a filial de faturamento
    lErro = Carrega_FilialFaturamento
    If lErro <> SUCESSO Then gError 119516
        
    'inicializa o grid DVV
    lErro = Inicializa_Grid_DVV(objGridDVV)
    If lErro <> SUCESSO Then gError 119517
    
    'inicializa o grid de analise
    lErro = Inicializa_Grid_Analise(objGridAnalise)
    If lErro <> SUCESSO Then gError 119518
        
    'zera as variaveis de alteração
    iQtdeAlterada = 0
    iProdutoAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iFilialFatAlterada = 0
    iAlterado = 0
    bAlterouTabPrincipal = False
        
    lErro = Chama_Tela_Nova_Instancia1("ComissoesCalcula", gobjTelaComissoes)
    If lErro <> SUCESSO Then gError 119518
    
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then gError 26481
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 119515 To 119518

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162624)

    End Select

    Exit Sub

End Sub

Private Sub Botao_Insumos_Click()

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoImprimir_Click()
'grava e imprime os dados do relatorio

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'carrega o gobj c/ os parametros da tela
    lErro = Move_Tela_Memoria
    If lErro <> SUCESSO Then gError 119519
     
    'chama a rotina de gravação da opção de impressão
    lErro = CF("RelMargContr_Grava", gobjMargContr)
    If lErro <> SUCESSO Then gError 119521
    
    'Executa o(s) Relatorio(s) de acordo com a selecao
    lErro = objRelatorio.ExecutarDireto("Analise de Margem de Contribuição", "", 0, "", "NNUMINTREL", gobjMargContr.lNumIntRel)
    If lErro <> SUCESSO Then gError 119522
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 119519 To 119522
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162625)

    End Select
    
    Exit Sub

End Sub

Private Function Move_GridAnalise_Memoria() As Long
'preenche as coleções com os dados da tela referentes ao grid DVV

Dim iLinha As Integer
Dim iColuna As Integer

On Error GoTo Erro_Move_GridAnalise_Memoria

    'para cada linha
    For iLinha = 1 To objGridAnalise.iLinhasExistentes
        
        'para cada coluna
        For iColuna = 2 To objGridAnalise.colColuna.Count - 1
            
            'atualiza o valor da celula
            gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1)).dValor = StrParaDbl(GridAnalise.TextMatrix(iLinha, iColuna))
        
        Next

    Next
    
    Move_GridAnalise_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridAnalise_Memoria:

    Move_GridAnalise_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162626)

    End Select
    
    Exit Function

End Function

Private Function Move_GridDVV_Memoria() As Long
'preenche as coleções com os dados referentes ao grid DVV

Dim iLinha As Integer
Dim iColuna As Integer

On Error GoTo Erro_Move_GridDVV_Memoria

    'para cada linha
    For iLinha = 1 To objGridDVV.iLinhasExistentes
        
        'para cada coluna
        For iColuna = 2 To objGridDVV.colColuna.Count - 1
            
            'atualiza o valor da celula
            gobjMargContr.colDVVLinCol(gobjMargContr.IndDVV(iLinha, iColuna - 1)).dValor = StrParaDbl(GridDVV.TextMatrix(iLinha, iColuna))
        
        Next

    Next

    Move_GridDVV_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridDVV_Memoria:

    Move_GridDVV_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162627)

    End Select
    
    Exit Function

End Function

Private Sub BotaoInsumos_Click()

Dim lErro As Long, objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoInsumos_Click

    If Len(Trim(Produto.ClipText)) <> 0 Then
    
        'formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 119536
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objKit.sProdutoRaiz = sProdutoFormatado
            Call Chama_Tela("MatPrim", objKit)
            
        End If
        
    End If
    
    Exit Sub
     
Erro_BotaoInsumos_Click:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162628)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela

Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_BotaoLimpar

    'verifica se alguma coisa foi escrito/mudado na tela
    If iAlterado = REGISTRO_ALTERADO Then
        
        'pergunta se deseja limpar a tela
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_LIMPAR_TELA")
        
        'se a resposta for não, sai da rotina
        If vbMsg = vbNo Then Exit Sub
    
    End If
    
    'limpa as 3 tabs
    Call Limpa_Tela_MargContr

    Exit Sub

Erro_BotaoLimpar:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162629)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_MargContr()
'limpa a tela toda

    'limpa a tela e os grids
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridDVV)
    Call Grid_Limpa(objGridAnalise)

    'limpas as combos e as labels
    Filial.Clear
    Descricao.Caption = ""
    LabelUM.Caption = ""
    FilialFaturamento.Text = ""

    'zera as variaveis de alteração
    iQtdeAlterada = 0
    iProdutoAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iFilialFatAlterada = 0
    iAlterado = 0

End Sub

Private Sub BotaoLimpaRecalculaColuna_Click()

Dim iLin As Integer, iColuna As Integer

    iColuna = GridAnalise.Col
    
    If iColuna <> 0 Then
    
        For iLin = 1 To objGridAnalise.iLinhasExistentes
        
            GridAnalise.TextMatrix(iLin, iColuna) = ""
        
        Next
        
        Call Analise_LimpaCelulasL1(0, GridAnalise.Col)
        
        Call Analise_RecalcularColuna(GridAnalise.Col, True)
    
    End If
    
End Sub

Private Sub BotaoRecalcularColuna_Click()
'rotina que recalcula a coluna corrente

Dim iColuna As Integer

On Error GoTo Erro_BotaoRecalcularColuna_Click

    iColuna = GridAnalise.Col
    
    If iColuna <> 0 Then
    
        Call Analise_LimpaCelulasL1(0, iColuna)
        
        Call Analise_RecalcularColuna(iColuna, False)
    
    End If
    
    Exit Sub

Erro_BotaoRecalcularColuna_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162630)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoRecalcularPlanilha_Click()
'rotina que limpa e recalcula a planilha

Dim iColuna As Integer, lErro As Long, objLinCol As ClassPlanMargContrLinCol, iLin As Integer

On Error GoTo Erro_BotaoRecalcularPlanilha

    For iColuna = 2 To 4
    
        For iLin = 1 To objGridDVV.iLinhasExistentes
            GridDVV.TextMatrix(iLin, iColuna) = ""
        Next
        
        lErro = DVV_RecalcularColuna(iColuna)
        If lErro <> SUCESSO Then gError 106722
        
    Next

    For iColuna = 2 To 9
    
        For iLin = 1 To objGridAnalise.iLinhasExistentes
            GridAnalise.TextMatrix(iLin, iColuna) = ""
        Next
        
'        'limpar a linha 1 se nao houver formula para ela
'        Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(1, iColuna - 1))
'        If Len(Trim(objLinCol.sFormula)) = "" Then GridAnalise.TextMatrix(1, iColuna) = ""
    
        Call Analise_LimpaCelulasL1(0, iColuna)
        
        lErro = Analise_RecalcularColuna(iColuna, True)
        If lErro <> SUCESSO Then gError 106723
        
    Next

    Exit Sub

Erro_BotaoRecalcularPlanilha:

    Select Case gErr
    
        Case 106722, 106723
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162631)

    End Select

    Exit Sub
    
End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Filial_Click()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub FilialFaturamento_Click()
    iFilialFatAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
    iQtdeAlterada = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
'verifica se a quantidade é valida

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    If iQtdeAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    'verifica se a quantidade foi preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

        'não pode ser valor negativo e nem 0
        lErro = Valor_Positivo_Critica_Double(Quantidade.Text)
        If lErro <> SUCESSO Then gError 119523

        'coloca a qntd formatada no controle
        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    iQtdeAlterada = 0
    
    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True
    
    Select Case gErr

        Case 119523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162632)
    
    End Select

    Exit Sub

End Sub

Private Sub TabelaPreco_Click()
    
    bAlterouTabPrincipal = True

End Sub

Private Sub TabOpcao_BeforeClick(Cancel As Integer)
'rotina que verifica se o 1º tab está preenchido

Dim lErro As Long

On Error GoTo Erro_TabOpcao_BeforeClick

    'se o tab atual é o de identificação do cliente
    If TabOpcao.SelectedItem.Index <> TAB_Identificacao Then Exit Sub

    'verifica se está faltando algum dado obrigatório no 1º tab.
    lErro = Valida_Tab_Identificacao
    If lErro <> SUCESSO Then gError 119524
    
    Exit Sub

Erro_TabOpcao_BeforeClick:

Cancel = True

    Select Case gErr
        
        Case 119524
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162633)

    End Select
    
    Exit Sub

End Sub

Private Sub TabOpcao_Click()
'sub que muda a posição das tabs

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If TabOpcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabOpcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de DVV
        If TabOpcao.SelectedItem.Index = TAB_DVV Or TabOpcao.SelectedItem.Index = TAB_Analise Then

            'preenche o grid de DVV
            lErro = Preenche_GridDVV
            If lErro <> SUCESSO Then gError 119529

            'preenche o grid de analise
            lErro = Preenche_GridAnalise
            If lErro <> SUCESSO Then gError 119530

        End If

        'Esconde o frame atual, mostra o novo
        FramePrincipal(giFrameAtual).Visible = False
        FramePrincipal(TabOpcao.SelectedItem.Index).Visible = True

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabOpcao.SelectedItem.Index
       
        If bAlterouTabPrincipal Then
            
            lErro = Move_Tela_Memoria1
            If lErro <> SUCESSO Then gError 119530
            
            lErro = CF("MargContr_CalculaComissoes", gcolComissoes, gobjMargContr, gdPrecoComissoes, gobjTelaComissoes, Cliente.Text, LabelUM.Caption)
            If lErro <> SUCESSO Then gError 119530
            
            Call BotaoRecalcularPlanilha_Click
            bAlterouTabPrincipal = False
        End If
       
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 119529, 119530
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162634)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria() As Long
'carrega o obj com os dados da tela para a impressão

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    'Verifica se os campos obrigatórios foram preenchidos
    lErro = Valida_Tab_Identificacao
    If lErro <> SUCESSO Then gError 119632
    
    lErro = Move_Tela_Memoria1()
    If lErro <> SUCESSO Then gError 106720
    
    'move os dados do grid de dvv p/ as collections
    lErro = Move_GridDVV_Memoria
    If lErro <> SUCESSO Then gError 119633
    
    'move os dados do grid de analise p/ as collections
    lErro = Move_GridAnalise_Memoria
    If lErro <> SUCESSO Then gError 119635
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 106720, 119632, 119633, 119635
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162635)

    End Select
    
    Exit Function

End Function

Private Function Valida_Tab_Identificacao() As Long
'faz a validação do 1º tab

On Error GoTo Erro_Valida_Tab_Identificacao

    'verifica se o cliente está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 119531

    'verifica se a filial do cliente está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 119532

    'verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 119533

    'verifica se a qntd está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then gError 119534

    'verifica se a filial de faturamento está preenchida
    If Len(Trim(FilialFaturamento.Text)) = 0 Then gError 119535
    
    If TabelaPreco.ListIndex = -1 Then gError 106860
    
    Valida_Tab_Identificacao = SUCESSO
    
    Exit Function
    
Erro_Valida_Tab_Identificacao:

    Valida_Tab_Identificacao = gErr
    
    Select Case gErr
        
        Case 119531
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", gErr)
    
        Case 119532
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)
    
        Case 119533
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)
        
        Case 119534
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDO1", gErr)
            
        Case 119535
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FATURAMENTO_NAO_PREENCHIDA", gErr)
    
        Case 106860
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_INFORMADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162636)
    
    End Select
    
    Exit Function
    
End Function

Private Function Preenche_GridDVV() As Long
'preenche o grid de dvv c/ oq tem no bd

Dim objDVVLin As ClassDVVLin
Dim lErro As Long

On Error GoTo Erro_Preenche_GridDVV

    'verifica se ja foi feita a leitura
    If bGridDVVCarregado = False Then
    
        lErro = CF("MargContr_Le_DVV", gobjMargContr)
        If lErro <> SUCESSO Then gError 119537
        
        bGridDVVCarregado = True

    End If
    
    'preenche a descricao dos itens c/ oq estiver no bd (DVVLin)
    For Each objDVVLin In gobjMargContr.colDVVLin
        
        GridDVV.TextMatrix(objDVVLin.iLinha, iGrid_DVVDescricao_Col) = objDVVLin.sDescricao
        
    Next
    
    'coloca o nº de linhas existentes
    objGridDVV.iLinhasExistentes = gobjMargContr.colDVVLin.Count
            
    Preenche_GridDVV = SUCESSO
    
    Exit Function

Erro_Preenche_GridDVV:

    Preenche_GridDVV = gErr

    Select Case gErr
    
        Case 119537
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162637)

    End Select
    
    Exit Function

End Function

Private Function Preenche_GridAnalise() As Long
'preenche o grid de analise c/ oq tem no bd

Dim objPlanMargContrCol As ClassPlanMargContrCol
Dim objPlanMargContrLin As ClassPlanMargContrLin
Dim objPlanMargContrLinCol As ClassPlanMargContrLinCol
Dim lErro As Long

On Error GoTo Erro_Preenche_GridAnalise

    'verifica se ja foi lido e preenchido
    If bGridAnaliseCarregado = False Then

        lErro = CF("MargContr_Le_Analise", gobjMargContr)
        If lErro <> SUCESSO Then gError 119538
        
        bGridAnaliseCarregado = True

    End If

    'preenche o titulo das colunas lidos no bd (PlanMargContrCol)
    For Each objPlanMargContrCol In gobjMargContr.colPlanMargContrCol
        
        GridAnalise.TextMatrix(0, 1 + objPlanMargContrCol.iColuna) = objPlanMargContrCol.sTitulo
    
    Next
        
    'preenche a descricao dos itens (PlanMargContrLin)
    For Each objPlanMargContrLin In gobjMargContr.colPlanMargContrLin
        
        GridAnalise.TextMatrix(objPlanMargContrLin.iLinha, iGrid_Analise_Col) = objPlanMargContrLin.sDescricao
        
    Next
    
    'preenche o obj c/ as linha existentes
    objGridAnalise.iLinhasExistentes = gobjMargContr.colPlanMargContrLin.Count
    
    Preenche_GridAnalise = SUCESSO
    
    Exit Function

Erro_Preenche_GridAnalise:

    Preenche_GridAnalise = gErr

    Select Case gErr
    
        Case 119538
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162638)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_DVV(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Itens")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Simulação")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DVVDescricao.Name)
    objGridInt.colCampo.Add (DVVValor1.Name)
    objGridInt.colCampo.Add (DVVValor2.Name)
    objGridInt.colCampo.Add (DVVValor3.Name)
    
    'Colunas do Grid
    iGrid_DVVDescricao_Col = 1
    iGrid_DVVValor1_Col = 2
    iGrid_DVVValor2_Col = 3
    iGrid_DVVValor3_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridDVV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = GRIDDVV_MAX_LINHAS

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 17

    'Largura da primeira coluna
    GridDVV.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'não pode excluir linhas
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_DVV = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Analise(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Itens")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Praticado")
    objGridInt.colColuna.Add ("A")
    objGridInt.colColuna.Add ("B")
    objGridInt.colColuna.Add ("B 1")
    objGridInt.colColuna.Add ("B 2")
    objGridInt.colColuna.Add ("C")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AnaliseLinDescricao.Name)
    objGridInt.colCampo.Add (Valor1.Name)
    objGridInt.colCampo.Add (Valor2.Name)
    objGridInt.colCampo.Add (Valor3.Name)
    objGridInt.colCampo.Add (Valor4.Name)
    objGridInt.colCampo.Add (Valor5.Name)
    objGridInt.colCampo.Add (Valor6.Name)
    objGridInt.colCampo.Add (Valor7.Name)
    objGridInt.colCampo.Add (Valor8.Name)
    
    'Colunas do Grid
    iGrid_Analise_Col = 1
    iGrid_Valor1_Col = 2
    iGrid_Valor2_Col = 3
    iGrid_Valor3_Col = 4
    iGrid_Valor4_Col = 5
    iGrid_Valor5_Col = 6
    iGrid_Valor6_Col = 7
    iGrid_Valor7_Col = 8
    iGrid_Valor8_Col = 9
    
    'Grid do GridInterno
    objGridInt.objGrid = GridAnalise

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = GRIDANALISE_MAX_LINHAS
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 21

    'Largura da primeira coluna
    GridAnalise.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibe excluir linha do grid
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Analise = SUCESSO

    Exit Function

End Function

Private Sub Cliente_Change()
    iClienteAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Produto_Change()
    iProdutoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub FilialFaturamento_Change()
    iFilialFatAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim iCodFilial As Integer
Dim objcliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Cliente_Validate
      
    'se o cliente não foi alterado, sai da rotina
    If iClienteAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'se o cliente não foi preenchido, sai da rotina
    If Len(Trim(Cliente.Text)) = 0 Then
        'limpa a filial
        Filial.Clear
        Exit Sub
    End If

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 119539

    'coloca o nome reduzido do cliente na tela
    Cliente.Text = objcliente.sNomeReduzido

    TabelaPreco.ListIndex = -1
    If objcliente.iTabelaPreco > 0 Then
        TabelaPreco.Text = objcliente.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        'Se o Tipo estiver preenchido
        If objcliente.iTipo > 0 Then
            
            objTipoCliente.iCodigo = objcliente.iTipo
            
            'Lê o Tipo de Cliente
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 26514
            
            If objTipoCliente.iTabelaPreco > 0 Then
                TabelaPreco.Text = objTipoCliente.iTabelaPreco
                Call TabelaPreco_Validate(bSGECancelDummy)
            End If
        End If
        
    End If
    
    'busca no bd a relação de filiais referentes ao cliente
    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 119540
    
    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    
    'verifica se foi digitado nome ou cód. do cliente
    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
            
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    End If
    
    iClienteAlterado = 0
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 119539, 119540
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162639)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)
'verifica se a filial de cliente é valida

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    If iFilialCliAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'verifica se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 119542
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 119543

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'passa o nº preenchido como código
        objFilialCliente.iCodFilial = iCodigo

        'Tentativa de leitura da Filial com esse código no BD
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 119545

        'Não encontrou Filial no  BD
        If lErro = 12567 Then gError 119544

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If
        
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 119546

    iFilialCliAlterada = 0
    
    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 119542
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
    
        Case 119543, 119545

        Case 119546, 119544
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162640)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)
'verifica se o produto é válido

Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate

    'se o produto não foi alterado => sai da função
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'se o produto nao estiver preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then
        
        'limpa as labels de descrição e sai da rotina
        Descricao.Caption = ""
        LabelUM.Caption = ""
        Exit Sub
    
    End If

    'Critica o formato do codigo
    lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido, True)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 119547
            
    'lErro = 25041 => inexistente
    If lErro = 25041 Then gError 119548
        
    'exibe os dados do produto na tela
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    'exibe a descrição
    Descricao.Caption = objProduto.sDescricao
    
    'exibe a uni. de medida
    LabelUM.Caption = objProduto.sSiglaUMEstoque
    
    'zera a variavel de alteração do produto
    iProdutoAlterado = 0
        
    Exit Sub
    
Erro_Produto_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 119547
            'limpa a label da descricao e da un
            Descricao.Caption = ""
            LabelUM.Caption = ""
            
        Case 119548
           'Não encontrou Produto no BD e pergunta se deseja criar um novo
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            
            'se sim
            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)
            'senão
            Else
                'limpa a label da descricao e da un
                Descricao.Caption = ""
                LabelUM.Caption = ""
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162641)
            
    End Select

    Exit Sub

End Sub

Private Sub Vendedor_Change()
    iVendedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    bAlterouTabPrincipal = True
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
'verifica se o vendedor existe

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Vendedor

    'verifica se o vendedor foi alterado
    If iVendedorAlterado = 0 Then Exit Sub
    
    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then Exit Sub

    'Verifica se Vendedor está cadastrado no bd
    lErro = TP_Vendedor_Le(Vendedor, objVendedor)
    If lErro <> SUCESSO Then gError 119549
    
    'zera a variavel de alteração do vendedor
    iVendedorAlterado = 0

    Exit Sub

Erro_Saida_Celula_Vendedor:

    Cancel = True

    Select Case gErr

        Case 119549
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162642)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialFaturamento() As Long
'Carrega FilialFaturamento com as Filiais Empresas

Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialFaturamento

    'p/ cada filial na coleção
    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then
        
            'adiciona a filial na combobox
            FilialFaturamento.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialFaturamento.ItemData(FilialFaturamento.NewIndex) = objFiliais.iCodFilial
    
        End If
        
    Next

    Call Seleciona_FilialEmpresa
    
    Carrega_FilialFaturamento = SUCESSO

    Exit Function

Erro_Carrega_FilialFaturamento:

    Carrega_FilialFaturamento = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162643)

    End Select

    Exit Function

End Function

Public Sub FilialFaturamento_Validate(Cancel As Boolean)
'verifica se a filal é válida

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialFaturamento_Validate
    
    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialFaturamento.Text)) = 0 Then Exit Sub

    'verifica se ela foi modificada
    If iFilialFatAlterada <> REGISTRO_ALTERADO Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFaturamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 119551

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 119553

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 119554

    'zera a variavel de alteracao
    iFilialFatAlterada = 0
    
    Exit Sub

Erro_FilialFaturamento_Validate:

    Cancel = True

    Select Case gErr

        Case 119550, 119551

        Case 119553
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 119554
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialFaturamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162644)

    End Select

    Exit Sub

End Sub

Sub GridDVV_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDVV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridDVV, iAlterado)

    End If
    
End Sub

Private Sub GridDVV_GotFocus()
    Call Grid_Recebe_Foco(objGridDVV)
End Sub

Private Sub GridDVV_EnterCell()
    Call Grid_Entrada_Celula(objGridDVV, iAlterado)
End Sub

Private Sub GridDVV_LeaveCell()
    Call Saida_Celula(objGridDVV)
End Sub

Private Sub GridDVV_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDVV)
End Sub

Private Sub GridDVV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDVV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDVV, iAlterado)
    End If

End Sub

Private Sub GridDVV_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDVV)
End Sub

Private Sub GridDVV_RowColChange()
    Call Grid_RowColChange(objGridDVV)
End Sub

Private Sub GridDVV_Scroll()
    Call Grid_Scroll(objGridDVV)
End Sub

Private Sub GridAnalise_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAnalise, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridAnalise, iAlterado)

    End If
    
End Sub

Private Sub GridAnalise_GotFocus()
    Call Grid_Recebe_Foco(objGridAnalise)
End Sub

Private Sub GridAnalise_EnterCell()
    Call Grid_Entrada_Celula(objGridAnalise, iAlterado)
End Sub

Private Sub GridAnalise_LeaveCell()
    Call Saida_Celula(objGridAnalise)
End Sub

Private Sub GridAnalise_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridAnalise)
End Sub

Private Sub GridAnalise_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAnalise, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnalise, iAlterado)
    End If

End Sub

Sub GridAnalise_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAnalise)
End Sub

Private Sub GridAnalise_RowColChange()
    Call Grid_RowColChange(objGridAnalise)
End Sub

Private Sub GridAnalise_Scroll()
    Call Grid_Scroll(objGridAnalise)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    
    'aqui está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridDVV
            Case GridDVV.Name

                'executa a saida de celula do grid de dvv
                lErro = Saida_Celula_GridDVV(objGridInt)
                If lErro <> SUCESSO Then gError 119555

            'Se for o GridAnalise
            Case GridAnalise.Name

                'executa a saida de celula do grid de analise
                lErro = Saida_Celula_GridAnalise(objGridInt)
                If lErro <> SUCESSO Then gError 119556

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 119557

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 119555, 119556, 119557

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162645)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDVV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do griddvv que está deixando de ser a corrente

Dim lErro As Long, dValor3Ant As Double

On Error GoTo Erro_Saida_Celula_GridDVV

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_DVVValor1_Col
            
            'Faz a saída do controle valor1
            lErro = Saida_Celula_DVVValor1(objGridInt)
            If lErro <> SUCESSO Then gError 119558

        Case iGrid_DVVValor2_Col
            
            'Faz a saída do controle valor2
            lErro = Saida_Celula_DVVValor2(objGridInt)
            If lErro <> SUCESSO Then gError 119559
        
        Case iGrid_DVVValor3_Col
            
            dValor3Ant = StrParaDbl(GridDVV.TextMatrix(objGridDVV.iLinhasExistentes, 4))
            
            'Faz a saída do controle valor3
            lErro = Saida_Celula_DVVValor3(objGridInt)
            If lErro <> SUCESSO Then gError 119560
    
    End Select

    Call DVV_RecalcularColuna(objGridInt.objGrid.Col)
    
    If objGridInt.objGrid.Col = iGrid_DVVValor3_Col And Abs(dValor3Ant - StrParaDbl(GridDVV.TextMatrix(objGridDVV.iLinhasExistentes, 4))) > DELTA_VALORMONETARIO Then
        Call DVV_AlterarDVV3
    End If
    
    Saida_Celula_GridDVV = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDVV:

    Saida_Celula_GridDVV = gErr

    Select Case gErr

        Case 119558, 119559, 119560

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162646)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DVVValor1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor 1

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = DVVValor1

    'Se o controle estiver preenchido
    If Len(Trim(DVVValor1.Text)) > 0 Then
                
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119561

    Saida_Celula_DVVValor1 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_DVVValor1 = gErr

    Select Case gErr

        Case 119561
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162647)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DVVValor2(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor 2

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = DVVValor2

    'Se o controle estiver preenchido
    If Len(Trim(DVVValor2.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119562

    Saida_Celula_DVVValor2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_DVVValor2 = gErr

    Select Case gErr

        Case 119562
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162648)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DVVValor3(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor 3

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = DVVValor3

    'Se o controle estiver preenchido
    If Len(Trim(DVVValor3.Text)) > 0 Then
        
        '???Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119563

    Saida_Celula_DVVValor3 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_DVVValor3 = gErr

    Select Case gErr

        Case 119563
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162649)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridAnalise(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do gridanalise que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridAnalise

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Valor1_Col
            
            'Faz a saída do controle valor1
            lErro = Saida_Celula_AnaliseValor1(objGridInt)
            If lErro <> SUCESSO Then gError 119564

        Case iGrid_Valor2_Col
            
            'Faz a saída do controle valor2
            lErro = Saida_Celula_AnaliseValor2(objGridInt)
            If lErro <> SUCESSO Then gError 119564

        Case iGrid_Valor3_Col
            
            'Faz a saída do controle valor3
            lErro = Saida_Celula_AnaliseValor3(objGridInt)
            If lErro <> SUCESSO Then gError 119564

        Case iGrid_Valor4_Col
            
            'Faz a saída do controle valor4
            lErro = Saida_Celula_AnaliseValor4(objGridInt)
            If lErro <> SUCESSO Then gError 119564

        Case iGrid_Valor5_Col
            
            'Faz a saída do controle valor5
            lErro = Saida_Celula_AnaliseValor5(objGridInt)
            If lErro <> SUCESSO Then gError 119565
        
        Case iGrid_Valor6_Col
            
            'Faz a saída do controle valor6
            lErro = Saida_Celula_AnaliseValor6(objGridInt)
            If lErro <> SUCESSO Then gError 119566
    
        Case iGrid_Valor7_Col
        
            'Faz a saída do controle valor7
            lErro = Saida_Celula_AnaliseValor7(objGridInt)
            If lErro <> SUCESSO Then gError 119567
        
        Case iGrid_Valor8_Col
            
            'Faz a saída do controle valor8
            lErro = Saida_Celula_AnaliseValor8(objGridInt)
            If lErro <> SUCESSO Then gError 119568
    
    End Select

    'limpar todas as celulas da coluna corrente que façam parte do grupo L1
    Call Analise_LimpaCelulasL1(objGridInt.objGrid.Row, objGridInt.objGrid.Col)
    
    Call Analise_RecalcularColuna(objGridInt.objGrid.Col, False)
    
    Saida_Celula_GridAnalise = SUCESSO

    Exit Function

Erro_Saida_Celula_GridAnalise:

    Saida_Celula_GridAnalise = gErr

    Select Case gErr

        Case 119564 To 119568

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162650)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor4(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor4

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor4

    'Se o controle estiver preenchido
    If Len(Trim(Valor4.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119569

    Saida_Celula_AnaliseValor4 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor4 = gErr

    Select Case gErr

        Case 119569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162651)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor5(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor5

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor5

    'Se o controle estiver preenchido
    If Len(Trim(Valor5.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119570

    Saida_Celula_AnaliseValor5 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor5 = gErr

    Select Case gErr

        Case 119570
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162652)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor6(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor6

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor6

    'Se o controle estiver preenchido
    If Len(Trim(Valor6.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119571

    Saida_Celula_AnaliseValor6 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor6 = gErr

    Select Case gErr

        Case 119571
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162653)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor7(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor7

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor7

    'Se o controle estiver preenchido
    If Len(Trim(Valor7.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119572

    Saida_Celula_AnaliseValor7 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor7 = gErr

    Select Case gErr

        Case 119572
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162654)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor8(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor8

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor8

    'Se o controle estiver preenchido
    If Len(Trim(Valor8.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119573

    Saida_Celula_AnaliseValor8 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor8 = gErr

    Select Case gErr

        Case 119573
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162655)

    End Select

    Exit Function

End Function

Private Sub LabelCliente_Click()
'sub chamadora do browser de clientes

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'se o cliente estiver preenchido
    If Len(Trim(Cliente.Text)) <> 0 Then
        'Prenche o obj c/ o nomereduzido do Cliente
        objcliente.sNomeReduzido = Cliente.Text
    End If

    'chama a tela de clientes
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub LabelProduto_Click()
'sub chamadora do browser Produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 119574

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 119574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162656)

    End Select

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()
'sub que chama o browser de vendedores

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'se o vendedor estiver preenchido
    If Len(Trim(Vendedor.Text)) <> 0 Then
        'carrega o obj c/ o nomereduzido do vendedor
        objVendedor.sNomeReduzido = Vendedor.Text
    End If
    
    'Chama tela que lista todos os vendores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser cliente

Dim objcliente As ClassCliente

    Set objcliente = obj1

    'Preenche o controle com o cod. do cliente selecionado
    Cliente.Text = objcliente.lCodigo
    
    'dispara o validate do cliente
    Call Cliente_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Produto

Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche campo Produto
    Produto.PromptInclude = False
    Produto.Text = CStr(objProduto.sCodigo)
    Produto.PromptInclude = True
    Call Produto_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162657)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser vendedor

Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
    
    'Preenche o Vendedor c/ o nomereduzido
    Vendedor.Text = objVendedor.iCodigo

    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162658)

    End Select

    Exit Sub

End Sub

Private Sub DVVValor1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVValor2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVValor3_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVValor1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVValor2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVValor3_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVValor1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = DVVValor1
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DVVValor2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = DVVValor2
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DVVValor3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = DVVValor3
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor4_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor5_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor6_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor7_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor8_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor4_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor5_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor6_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor7_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor8_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor4_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor4
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor5_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor5
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor6_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor6
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor7_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor7
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor8_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor8
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        End If
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Sub Form_Unload(Cancel As Integer)
'finaliza os objs

    'objs dos browser
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
   
    'objs do grid
    Set objGridAnalise = Nothing
    Set objGridDVV = Nothing
   
    'obj global
    Set gobjMargContr = Nothing
   
    If Not gobjTelaComissoes Is Nothing Then
    
        gobjTelaComissoes.Unload gobjTelaComissoes
        Set gobjTelaComissoes = Nothing
        
    End If
    
    Set gcolComissoes = Nothing
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Análise de Margem de Contribuição"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MargContr"
    
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

Public Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
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

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LabelQuantidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQuantidade, Source, X, Y)
End Sub

Private Sub LabelQuantidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQuantidade, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFat_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFat, Source, X, Y)
End Sub

Private Sub LabelFilialFat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFat, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedor, Source, X, Y)
End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedor, Button, Shift, X, Y)
End Sub

'Public Function RelMargContr_Grava(ByVal objMargContr As ClassMargContr) As Long
''grava os dados do relatorio nas tabelas temporárias para ser impresso
'
'Dim lErro As Long
'Dim lNumIntRel As Long
'Dim lTransacao As Long
'
'On Error GoTo Erro_RelMargContr_Grava
'
'    'Inicia a Transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 119612
'
'    'obtem o nº automatico
'    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
'    If lErro <> SUCESSO Then gError 119613
'
'    'carrega o obj c/ o nº automatico
'    objMargContr.lNumIntRel = lNumIntRel
'
'    'grava na tabela RelMargContr
'    lErro = CF("RelMargContr_Grava_EmTrans", objMargContr)
'    If lErro <> SUCESSO Then gError 119614
'
'    'grava na tabela RelMargContrCol
'    lErro = CF("RelMargContrCol_Grava_EmTrans", objMargContr)
'    If lErro <> SUCESSO Then gError 119615
'
'    'grava na tabela RelMargContrLin
'    lErro = CF("RelMargContrLin_Grava_EmTrans", objMargContr)
'    If lErro <> SUCESSO Then gError 119616
'
'    'grava na tabela RelMargContrLinCol
'    lErro = CF("RelMargContrLinCol_Grava_EmTrans", objMargContr)
'    If lErro <> SUCESSO Then gError 119617
'
'    'confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 119622
'
'    RelMargContr_Grava = SUCESSO
'
'    Exit Function
'
'Erro_RelMargContr_Grava:
'
'    RelMargContr_Grava = gErr
'
'    Select Case gErr
'
'        Case 119612
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_ABRIR", gErr)
'
'        Case 119613 To 119617
'
'        Case 119622
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162659)
'
'    End Select
'
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function
'
'Public Function RelMargContr_Grava_EmTrans(ByVal objMargContr As ClassMargContr) As Long
''grava na tabela RelMargContr com o objMargContr passado como parametro
'
'Dim lErro As Long
'Dim lComando As Long
'
'On Error GoTo Erro_RelMargContr_Grava_EmTrans
'
'    'abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 119618
'
'    'grava na tabela RelMargContr
'    lErro = Comando_Executar(lComando, "INSERT INTO RelMargContr(NumIntRel, CodCliente, CodFilial, CodVendedor, FilialFaturamento, Produto, Quantidade) VALUES (?,?,?,?,?,?,?)", objMargContr.lNumIntRel, objMargContr.lCodCliente, objMargContr.iCodFilial, objMargContr.iCodVendedor, objMargContr.iFilialFaturamento, objMargContr.sProduto, objMargContr.dQuantidade)
'    If lErro <> AD_SQL_SUCESSO Then gError 119619
'
'    'fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RelMargContr_Grava_EmTrans = SUCESSO
'
'    Exit Function
'
'Erro_RelMargContr_Grava_EmTrans:
'
'    RelMargContr_Grava_EmTrans = gErr
'
'    Select Case gErr
'
'        Case 119618
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 119619
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMARG", gErr)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Function RelMargContrLin_Grava_EmTrans(ByVal objMargContr As ClassMargContr) As Long
''grava na tabela RelMargContrLin a partir da col. passada no obj
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objPlanMargContrLin As ClassPlanMargContrLin, objDVVLin As ClassDVVLin
'
'On Error GoTo Erro_RelMargContrLin_Grava_EmTrans
'
'    'abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 119620
'
'    'Para cada objRelMargContrLin na coleção
'    For Each objPlanMargContrLin In objMargContr.colPlanMargContrLin
'
'        'insere um novo registro
'        lErro = Comando_Executar(lComando, "INSERT INTO RelMargContrLin(NumIntRel, TipoReg, Linha, Descricao, Formato) VALUES (?,?,?,?,?)", objMargContr.lNumIntRel, MARGCONTR_GRIDANALISE, objPlanMargContrLin.iLinha, objPlanMargContrLin.sDescricao, objPlanMargContrLin.iFormato)
'        If lErro <> AD_SQL_SUCESSO Then gError 119621
'
'    Next
'
'    'Para cada objRelMargContrLin na coleção
'    For Each objDVVLin In objMargContr.colDVVLin
'
'        'insere um novo registro
'        lErro = Comando_Executar(lComando, "INSERT INTO RelMargContrLin(NumIntRel, TipoReg, Linha, Descricao, Formato) VALUES (?,?,?,?,?)", objMargContr.lNumIntRel, MARGCONTR_GRIDDVV, objDVVLin.iLinha, objDVVLin.sDescricao, 4444) '??? trocar 4444 por cte devida
'        If lErro <> AD_SQL_SUCESSO Then gError 119621
'
'    Next
'
'    'fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RelMargContrLin_Grava_EmTrans = SUCESSO
'
'    Exit Function
'
'Erro_RelMargContrLin_Grava_EmTrans:
'
'    RelMargContrLin_Grava_EmTrans = gErr
'
'    Select Case gErr
'
'        Case 119620
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 119621
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMARGLIN", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162660)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Function RelMargContrLinCol_Grava_EmTrans(ByVal objMargContr As ClassMargContr) As Long
''grava na tabela RelMargContrLinCol a partir da col. passada no obj
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objPlanMargContrLinCol As ClassPlanMargContrLinCol, objDVVLinCol As ClassDVVLinCol
'
'On Error GoTo Erro_RelMargContrLinCol_Grava_EmTrans
'
'    'abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 119623
'
'    'Para cada objPlanMargContrLinCol na coleção
'    For Each objPlanMargContrLinCol In objMargContr.colPlanMargContrLinCol
'
'        'insere um novo registro
'        lErro = Comando_Executar(lComando, "INSERT INTO RelMargContrLinCol(NumIntRel, TipoReg, Linha, Coluna, Valor) VALUES (?,?,?,?,?)", objMargContr.lNumIntRel, MARGCONTR_GRIDANALISE, objPlanMargContrLinCol.iLinha, objPlanMargContrLinCol.iColuna, objPlanMargContrLinCol.dValor)
'        If lErro <> AD_SQL_SUCESSO Then gError 119624
'
'    Next
'
'    'Para cada objDVVLinCol na coleção
'    For Each objDVVLinCol In objMargContr.colDVVLinCol
'
'        'insere um novo registro
'        lErro = Comando_Executar(lComando, "INSERT INTO RelMargContrLinCol(NumIntRel, TipoReg, Linha, Coluna, Valor) VALUES (?,?,?,?,?)", objMargContr.lNumIntRel, MARGCONTR_GRIDANALISE, objDVVLinCol.iLinha, objDVVLinCol.iColuna, objDVVLinCol.dValor)
'        If lErro <> AD_SQL_SUCESSO Then gError 119624
'
'    Next
'
'    'fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RelMargContrLinCol_Grava_EmTrans = SUCESSO
'
'    Exit Function
'
'Erro_RelMargContrLinCol_Grava_EmTrans:
'
'    RelMargContrLinCol_Grava_EmTrans = gErr
'
'    Select Case gErr
'
'        Case 119623
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 119624
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMARGLINCOL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162661)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Function RelMargContrCol_Grava_EmTrans(ByVal objMargContr As ClassMargContr) As Long
''grava na tabela RelMargContrCol a partir da col. passado no obj
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objPlanMargContrCol As ClassPlanMargContrCol
'
'On Error GoTo Erro_RelMargContrCol_Grava_EmTrans
'
'    'abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 119626
'
'    'Para cada objPlanMargContrCol na coleção
'    For Each objPlanMargContrCol In objMargContr.colPlanMargContrCol
'
'        'insere um novo registro
'        lErro = Comando_Executar(lComando, "INSERT INTO RelMargContrCol(NumIntRel, TipoReg, Coluna, Titulo) VALUES (?,?,?,?)", objMargContr.lNumIntRel, MARGCONTR_GRIDANALISE, objPlanMargContrCol.iColuna, objPlanMargContrCol.sTitulo)
'        If lErro <> AD_SQL_SUCESSO Then gError 119627
'
'    Next
'
'    'fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RelMargContrCol_Grava_EmTrans = SUCESSO
'
'    Exit Function
'
'Erro_RelMargContrCol_Grava_EmTrans:
'
'    RelMargContrCol_Grava_EmTrans = gErr
'
'    Select Case gErr
'
'        Case 119626
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 119627
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMARGCOL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162662)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'??? Fim - Remover o trecho acima p/ RotinasFatGrava

Private Sub Seleciona_FilialEmpresa()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialFaturamento As Integer

On Error GoTo Erro_Seleciona_FilialEmpresa
    
    iFilialFaturamento = IIf(gobjFAT.iFilialFaturamento <> 0, gobjFAT.iFilialFaturamento, giFilialEmpresa)
    
    If iFilialFaturamento <> EMPRESA_TODA Then
        'seleciona a filial de faturamento na combo
        For iIndice = 0 To FilialFaturamento.ListCount - 1

            If FilialFaturamento.ItemData(iIndice) = iFilialFaturamento Then

                FilialFaturamento.ListIndex = iIndice
                Exit For

            End If
        Next

    Else
        FilialFaturamento.ListIndex = 0
    End If

    Exit Sub

Erro_Seleciona_FilialEmpresa:

    Select Case gErr

        Case 51139

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162663)

    End Select

    Exit Sub

End Sub

Function DVV_RecalcularColuna(ByVal iColuna As Integer) As Long

Dim lErro As Long, iLinha As Integer, objDVVLinCol As ClassDVVLinCol
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, dtDataCF As Date

On Error GoTo Erro_DVV_RecalcularColuna

    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", gobjMargContr.iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError 106858
    
    For iLinha = 1 To gobjMargContr.colDVVLin.Count
            
        Set objDVVLinCol = gobjMargContr.colDVVLinCol(gobjMargContr.IndDVV(iLinha, iColuna - 1))
        
        If Len(Trim(objDVVLinCol.sFormula)) = 0 Then
        
            'pega conteudo da tela, da propria coluna
            sFormula = GridDVV.TextMatrix(iLinha, iColuna)
            
            If Len(Trim(sFormula)) = 0 Then
            
                'pega conteudo da tela, da coluna vizinha à esquerda
                sFormula = GridDVV.TextMatrix(iLinha, iColuna - 1)
                
            End If
            
            If Len(Trim(sFormula)) <> 0 Then sFormula = CStr(StrParaDbl(sFormula) / 100)
            Call TrocaPontoVirgula(sFormula)
            
        Else
        
            sFormula = objDVVLinCol.sFormula
            
        End If
        
        Set objPlanilhas = New ClassPlanilhas
        
        With objPlanilhas
            .iTipoPlanilha = PLANILHA_TIPO_DVV
            .iFilialEmpresa = gobjMargContr.iFilialFaturamento
            .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
            .iLinha = iLinha
            .sExpressao = sFormula
        End With
        
        colPlanilhas.Add objPlanilhas
    
    Next
    
    With objContexto
        .iFilialFaturamento = gobjMargContr.iFilialFaturamento
        .sProduto = gobjMargContr.sProduto
        .dQuantidade = gobjMargContr.dQuantidade
        .iFilialCli = gobjMargContr.iCodFilial
        .lCliente = gobjMargContr.lCodCliente
        .iVendedor = gobjMargContr.iCodVendedor
        .iTabelaPreco = gobjMargContr.iTabelaPreco
        .iAno = Year(gdtDataAtual)
        .dtDataCustoFixo = dtDataCF
        .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR
        Set .colComissoes = gcolComissoes
        .dPrecoPraticado = gdPrecoComissoes
        .sNomeRedCliente = Cliente.Text
        .sUM = LabelUM.Caption
    End With
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 106721

    For Each objPlanilhas In colPlanilhas
    
        GridDVV.TextMatrix(objPlanilhas.iLinha, iColuna) = Format(objPlanilhas.dValor * 100, "###,##0.00###")

    Next
    
    DVV_RecalcularColuna = SUCESSO
     
    Exit Function
    
Erro_DVV_RecalcularColuna:

    DVV_RecalcularColuna = gErr
     
    Select Case gErr
          
        Case 106721
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162664)
     
    End Select
     
    Exit Function

End Function

Function Analise_RecalcularColuna(ByVal iColuna As Integer, ByVal bLimpando As Boolean) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dtDataCF As Date, dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double

On Error GoTo Erro_Analise_RecalcularColuna

    lErro = CF("CalcMP_ObterValores", giFilialEmpresa, dTaxaDescPadrao, dTaxaValFut, dDiasValFut)
    If lErro <> SUCESSO Then gError 106885

    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", gobjMargContr.iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError 106858
    
    'se a linha 1 nao está preenchida
    If Len(Trim(GridAnalise.TextMatrix(1, iColuna))) = 0 Then
    
        With objContexto2
            .iFilialFaturamento = gobjMargContr.iFilialFaturamento
            .sProduto = gobjMargContr.sProduto
            .dQuantidade = gobjMargContr.dQuantidade
            .iFilialCli = gobjMargContr.iCodFilial
            .lCliente = gobjMargContr.lCodCliente
            .iVendedor = gobjMargContr.iCodVendedor
            .iTabelaPreco = gobjMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR
            Set .colComissoes = gcolComissoes
            .dPrecoPraticado = gdPrecoComissoes
            .dTaxaDescPadrao = dTaxaDescPadrao
            .dTaxaValFut = dTaxaValFut
            .dDiasValFut = dDiasValFut
        End With
        
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objContexto2)
        
        For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
        
            Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
            
            If bLimpando Then
                sFormula = objLinCol.sFormula
                If iLinha = 1 And Len(Trim(sFormula)) <> 0 Then Exit For
                If Len(Trim(objLin.sFormulaL1)) = 0 And Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            Else
                sFormula = GridAnalise.TextMatrix(iLinha, iColuna)
                If Len(Trim(sFormula)) <> 0 And objLin.iFormato = GRID_FORMATO_PERCENTAGEM Then sFormula = CStr(StrParaDbl(sFormula) / 100)
                Call TrocaPontoVirgula(sFormula)
            End If
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = gobjMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas2.Add objPlanilhas
            
            'se possui formula p/calculo da linha 1 e o valor está preenchido...
            If Len(Trim(objLin.sFormulaL1)) <> 0 And ((bLimpando And Len(Trim(objLinCol.sFormula)) <> 0) Or (bLimpando = False And Len(Trim(GridAnalise.TextMatrix(iLinha, iColuna))) <> 0)) Then
            
                Set objPlanilhas = New ClassPlanilhas
                
                With objPlanilhas
                    .iTipoPlanilha = PLANILHA_TIPO_TODOS
                    .iFilialEmpresa = gobjMargContr.iFilialFaturamento
                    .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                    .iLinha = iLinha + 1
                    .sExpressao = objLin.sFormulaL1
                End With
                
                colPlanilhas2.Add objPlanilhas
        
                'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
                lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas2, dValor, objContexto2)
                If lErro <> SUCESSO Then gError 106721

                GridAnalise.TextMatrix(1, iColuna) = CStr(dValor)
                                
                Exit For
                
            End If
            
        Next
        
    End If
    
    Call Analise_LimpaCelulasNaoEdit(iColuna)

    'guarda os resultados do grid dvv para todas as colunas
    Call DVVTotal_GuardaValores(objContexto)
    
    For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
    
        Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
        Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
            
        If ((iLinha = 1 And Len(Trim(objLinCol.sFormula)) = 0 And Len(Trim(objLin.sFormulaGeral)) = 0) Or (bLimpando = False And iLinha <> 1 And Len(Trim(objLin.sFormulaL1)) = 0) And objLin.iEditavel <> 0) Then
        
            'pega conteudo da tela, da propria celula
            sFormula = GridAnalise.TextMatrix(objLinCol.iLinha, iColuna)
            If Len(Trim(sFormula)) <> 0 And objLin.iFormato = GRID_FORMATO_PERCENTAGEM Then sFormula = CStr(StrParaDbl(sFormula) / 100)
            Call TrocaPontoVirgula(sFormula)
            
        Else
        
            sFormula = objLinCol.sFormula
            If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
        
        End If
        
        If Len(Trim(sFormula)) = 0 Then sFormula = "0"
        
        Set objPlanilhas = New ClassPlanilhas
        
        With objPlanilhas
            .iTipoPlanilha = PLANILHA_TIPO_TODOS
            .iFilialEmpresa = gobjMargContr.iFilialFaturamento
            .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
            .iLinha = iLinha
            .sExpressao = sFormula
        End With
        
        colPlanilhas.Add objPlanilhas
    
    Next
    
    With objContexto
        .iFilialFaturamento = gobjMargContr.iFilialFaturamento
        .sProduto = gobjMargContr.sProduto
        .dQuantidade = gobjMargContr.dQuantidade
        .iFilialCli = gobjMargContr.iCodFilial
        .lCliente = gobjMargContr.lCodCliente
        .iVendedor = gobjMargContr.iCodVendedor
        .iTabelaPreco = gobjMargContr.iTabelaPreco
        .iAno = Year(gdtDataAtual)
        .dtDataCustoFixo = dtDataCF
        .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR
        Set .colComissoes = gcolComissoes
        .dPrecoPraticado = gdPrecoComissoes
        .dTaxaDescPadrao = dTaxaDescPadrao
        .dTaxaValFut = dTaxaValFut
        .dDiasValFut = dDiasValFut
    End With
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 106721

    For Each objPlanilhas In colPlanilhas
    
        Set objLin = gobjMargContr.colPlanMargContrLin(objPlanilhas.iLinha)
        If objLin.iFormato = GRID_FORMATO_PERCENTAGEM Then
            GridAnalise.TextMatrix(objPlanilhas.iLinha, iColuna) = Format(objPlanilhas.dValor * 100, "###,##0.00###")
        Else
            GridAnalise.TextMatrix(objPlanilhas.iLinha, iColuna) = Format(objPlanilhas.dValor, "###,##0.00###")
        End If

    Next
    
    Analise_RecalcularColuna = SUCESSO
     
    Exit Function
    
Erro_Analise_RecalcularColuna:

    Analise_RecalcularColuna = gErr
     
    Select Case gErr
          
        Case 106721
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162665)
     
    End Select
     
    Exit Function

End Function

Private Function Move_Tela_Memoria1() As Long
'carrega o obj com os dados da tela para a impressão

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Move_Tela_Memoria1
    
    'carrega o objcliente p/ passar como parametro
    objcliente.sNomeReduzido = Trim(Cliente.Text)
    
    'busca o código do cliente a apartir do nomered
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 119629
    
    'se o cliente nao foi encontrado ==> erro
    If lErro = 12348 Then gError 119631
    
    'guarda no obj o cód. do cliente
    gobjMargContr.lCodCliente = objcliente.lCodigo
    
    'carrega o gobj c/ a filial do cliente
    gobjMargContr.iCodFilial = Codigo_Extrai(Filial.Text)
    
    'formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 119536

    'carrega o gobj c/ o produto formatado
    gobjMargContr.sProduto = sProdutoFormatado

    'carrega o gobj c/ a qntd
    gobjMargContr.dQuantidade = StrParaDbl(Quantidade.Text)
    
    'carrega o gobj c/ o cód. da filial de faturamento
    gobjMargContr.iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    
    'se o vendedor estiver preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        'carrega o objvendedor p/ ser passado como parametro
        objVendedor.sNomeReduzido = Trim(Vendedor.Text)
                
        'busca o cód do vendedor a partir do nomered
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 119630
        
        'se o vendedor não foi encontrado, ==> erro
        If lErro = 25008 Then gError 119634
        
        'preenche o gobj c/ o cód do vendedor
        gobjMargContr.iCodVendedor = CStr(objVendedor.iCodigo)
        
    End If
    
    gobjMargContr.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    
    Move_Tela_Memoria1 = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria1:

    Move_Tela_Memoria1 = gErr

    Select Case gErr
    
        Case 119536, 119629, 119630
    
        Case 119631
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)
    
        Case 119634
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, Vendedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162666)

    End Select
    
    Exit Function

End Function

Private Sub TrocaPontoVirgula(sNumero As String)
Dim iTam As Integer, sResult As String, i As Integer, sCaracter As String

    For i = 1 To Len(sNumero)
    
        sCaracter = Mid(sNumero, i, 1)
        Select Case sCaracter
        
            Case ","
                sCaracter = "."
                
            Case "."
                sCaracter = ""
            
        End Select
        
        sResult = sResult & sCaracter
        
    Next
        
    sNumero = sResult

End Sub

Sub Analise_LimpaCelulasL1(ByVal iLinha As Integer, ByVal iColuna As Integer)
'limpa as celulas da coluna informada que pertencam ao grupo L1, a menos da linha informada

Dim lErro As Long, iLin As Integer, objLin As ClassPlanMargContrLin

On Error GoTo Erro_Analise_LimpaCelulasL1

    For iLin = 2 To objGridAnalise.iLinhasExistentes
    
        Set objLin = gobjMargContr.colPlanMargContrLin(iLin)
        
        If (Len(Trim(objLin.sFormulaL1)) <> 0) Then
        
            If iLin <> iLinha Then
                GridAnalise.TextMatrix(iLin, iColuna) = ""
            Else
                GridAnalise.TextMatrix(1, iColuna) = ""
            End If
            
        End If
        
    Next
    
    Exit Sub
     
Erro_Analise_LimpaCelulasL1:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162667)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
'habilita / desabilita o campo produto

Dim lErro As Long
        
On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case "Valor1", "Valor2", "Valor3", "Valor4", "Valor5", "Valor6", "Valor7", "Valor8"
            If iLinha < 1 Or iLinha > gobjMargContr.colPlanMargContrLin.Count Then
                objControl.Enabled = False
            Else
                If gobjMargContr.colPlanMargContrLin.Item(iLinha).iEditavel = 0 Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
            End If
            
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 116492
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162668)
    
    End Select
    
    Exit Sub
    
End Sub

Sub Analise_LimpaCelulasNaoEdit(ByVal iColuna As Integer)
'limpa as celulas da coluna informada que NAO pertencam ao grupo L1 e nao sejam editaveis

Dim lErro As Long, iLin As Integer, objLin As ClassPlanMargContrLin

On Error GoTo Erro_Analise_LimpaCelulasNaoEdit

    For iLin = 2 To objGridAnalise.iLinhasExistentes
    
        Set objLin = gobjMargContr.colPlanMargContrLin(iLin)
        
        If objLin.iEditavel = 0 And (Len(Trim(objLin.sFormulaL1)) = 0) Then
        
            GridAnalise.TextMatrix(iLin, iColuna) = ""
            
        End If
        
    Next
    
    Exit Sub
     
Erro_Analise_LimpaCelulasNaoEdit:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162669)
     
    End Select
     
    Exit Sub

End Sub

'Private Function CalculaComissoes() As Long
'
'Dim lErro As Long
'Dim objItemPedido As New ClassItemPedido, dPrecoUnitario As Double
'Dim objItemNF As New ClassItemNF
'
'On Error GoTo Erro_CalculaComissoes
'
'    Set gcolComissoes = New Collection
'
'    objItemPedido.iFilialEmpresa = gobjMargContr.iFilialFaturamento
'    objItemPedido.sProduto = gobjMargContr.sProduto
'    lErro = CF("ClienteFilial_Le_UltimoItemPedido", objItemPedido, gobjMargContr.lCodCliente, gobjMargContr.iCodFilial)
'    If lErro <> SUCESSO And lErro <> 94412 Then gError 106677
'
'    'Se encontrou o Item de pedido
'    If lErro = SUCESSO Then
'
'        lErro = CF("Produto_ConvPrecoUMAnalise", objItemPedido.sProduto, objItemPedido.sUnidadeMed, objItemPedido.dPrecoUnitario, dPrecoUnitario)
'        If lErro <> SUCESSO Then gError 130002
'
'    Else
'
'        objItemNF.sProduto = gobjMargContr.sProduto
'        lErro = CF("ClienteFilial_Le_UltimoItemNFVenda", objItemNF, gobjMargContr.iFilialFaturamento, gobjMargContr.lCodCliente, gobjMargContr.iCodFilial)
'        If lErro = SUCESSO Then
'
'            lErro = CF("Produto_ConvPrecoUMAnalise", objItemNF.sProduto, objItemNF.sUnidadeMed, objItemNF.dPrecoUnitario, dPrecoUnitario)
'            If lErro <> SUCESSO Then gError 130001
'
'        End If
'
'    End If
'
'    If dPrecoUnitario = 0 Then dPrecoUnitario = 1
'
'    gdPrecoComissoes = dPrecoUnitario
'
'    lErro = gobjTelaComissoes.Trata_Parametros(gobjMargContr.iFilialFaturamento, Cliente.Text, gobjMargContr.iCodFilial, gobjMargContr.sProduto, gobjMargContr.dQuantidade, LabelUM.Caption, dPrecoUnitario, gcolComissoes)
'    If lErro <> SUCESSO Then gError 106861
'
'    CalculaComissoes = SUCESSO
'
'    Exit Function
'
'Erro_CalculaComissoes:
'
'    CalculaComissoes = gErr
'
'    Select Case gErr
'
'        Case 130002, 130003
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162670)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 26482

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next
    
    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 26482

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162671)

    End Select

    Exit Function

End Function

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26538

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 26539

        If lErro <> SUCESSO Then gError 26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

''        lErro = Trata_TabelaPreco()
''        If lErro <> SUCESSO Then gError 30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

    Case 26538, 26539, 30527


    Case 26540  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        Else
            'Segura o foco

        End If

    Case 26541

        lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162672)

    End Select

    Exit Sub

End Sub

Sub DVVTotal_GuardaValores(objContexto As ClassContextoPlan)

Dim iCol As Integer, objMnemonicoValor As ClassMnemonicoValor

    'guarda os resultados do grid dvv para todas as colunas
    For iCol = 2 To 4
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        Set objMnemonicoValor.colValor = New Collection
        
        objMnemonicoValor.sMnemonico = "DVVTotal"
        objMnemonicoValor.vParam(1) = CDbl(iCol - 1)
    
        If objGridDVV.iLinhasExistentes > 0 Then
            objMnemonicoValor.colValor.Add StrParaDbl(GridDVV.TextMatrix(objGridDVV.iLinhasExistentes, iCol) / 100)
        Else
            objMnemonicoValor.colValor.Add CDbl(0)
        End If
        
        objContexto.colMnemonicoValor.Add objMnemonicoValor
        
    Next

End Sub

Sub DVV_AlterarDVV3()
'forçar o recalculo das colunas que utilizam DVVTotal(3) em suas formulas

Dim iLinha As Integer, iColuna As Integer, sFormula As String
Dim objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin

    For iColuna = 2 To 9
    
        For iLinha = 2 To gobjMargContr.colPlanMargContrLin.Count
        
            Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
            
            sFormula = objLinCol.sFormula
            If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            If InStr(sFormula, "DVVTotal(3)") <> 0 Then
                
                GridAnalise.TextMatrix(1, iColuna) = ""
                GridAnalise.TextMatrix(iLinha, iColuna) = GridDVV.TextMatrix(objGridDVV.iLinhasExistentes, 4)
                Call Analise_RecalcularColuna(iColuna, False)
                Exit For
                
            End If
            
        Next
                
    Next
    
End Sub

Private Function Saida_Celula_AnaliseValor3(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor3

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor3

    'Se o controle estiver preenchido
    If Len(Trim(Valor3.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119569

    Saida_Celula_AnaliseValor3 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor3 = gErr

    Select Case gErr

        Case 119569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162651)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor2(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor2

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor2

    'Se o controle estiver preenchido
    If Len(Trim(Valor2.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119569

    Saida_Celula_AnaliseValor2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor2 = gErr

    Select Case gErr

        Case 119569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162651)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AnaliseValor1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor1

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = Valor1

    'Se o controle estiver preenchido
    If Len(Trim(Valor1.Text)) > 0 Then
        
        '??? Jones, fazer tratamento
        
    End If
            
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119569

    Saida_Celula_AnaliseValor1 = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_AnaliseValor1 = gErr

    Select Case gErr

        Case 119569
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162651)

    End Select

    Exit Function

End Function

Private Sub Valor1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor1
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor2
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor3_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnalise)
End Sub

Private Sub Valor3_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnalise)
End Sub

Private Sub Valor3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDVV.objControle = Valor3
    lErro = Grid_Campo_Libera_Foco(objGridAnalise)
    If lErro <> SUCESSO Then Cancel = True

End Sub


