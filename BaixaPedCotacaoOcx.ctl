VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPedCotacaoOcx 
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
      DragMode        =   1  'Automatic
      Height          =   8175
      Index           =   2
      Left            =   255
      TabIndex        =   36
      Top             =   795
      Width           =   16185
      Begin VB.TextBox Fornecedor 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   10110
         TabIndex        =   12
         Text            =   "Fornecedor"
         Top             =   3705
         Width           =   3375
      End
      Begin VB.CheckBox Baixa 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   735
         TabIndex        =   10
         Top             =   1110
         Width           =   870
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   885
         Left            =   2805
         Picture         =   "BaixaPedCotacaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7230
         Width           =   1665
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   885
         Left            =   780
         Picture         =   "BaixaPedCotacaoOcx.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7230
         Width           =   1665
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6525
         TabIndex        =   15
         Text            =   "Data Emissão "
         Top             =   1065
         Width           =   1410
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Editar Pedido de Cotação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   6630
         Picture         =   "BaixaPedCotacaoOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   7230
         Width           =   1665
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "BaixaPedCotacaoOcx.ctx":2E7A
         Left            =   1965
         List            =   "BaixaPedCotacaoOcx.ctx":2E8A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   315
         Width           =   2865
      End
      Begin VB.TextBox DataValidade 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7680
         TabIndex        =   16
         Text            =   "Data Validade"
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3705
         TabIndex        =   13
         Text            =   "Filial Fornecedor"
         Top             =   1095
         Width           =   1590
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1650
         TabIndex        =   11
         Text            =   "Pedido"
         Top             =   1110
         Width           =   975
      End
      Begin VB.CommandButton BotaoBaixa 
         Caption         =   "Baixar Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6480
         Picture         =   "BaixaPedCotacaoOcx.ctx":2EC6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   60
         Width           =   1830
      End
      Begin VB.TextBox Data 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   5340
         TabIndex        =   14
         Text            =   "Data"
         Top             =   1080
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedCotacao 
         Height          =   6240
         Left            =   570
         TabIndex        =   9
         Top             =   825
         Width           =   13230
         _ExtentX        =   23336
         _ExtentY        =   11007
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
         Left            =   615
         TabIndex        =   37
         Top             =   360
         Width           =   1410
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   15450
      Picture         =   "BaixaPedCotacaoOcx.ctx":302C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Fechar"
      Top             =   165
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8280
      Index           =   1
      Left            =   285
      TabIndex        =   22
      Top             =   735
      Width           =   16290
      Begin VB.Frame Frame7 
         Caption         =   "Exibe Pedidos"
         Height          =   4500
         Left            =   450
         TabIndex        =   23
         Top             =   330
         Width           =   7440
         Begin VB.Frame Frame11 
            Caption         =   "Data"
            Height          =   930
            Left            =   570
            TabIndex        =   31
            Top             =   2310
            Width           =   6420
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   1665
               TabIndex        =   4
               Top             =   345
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataEmissaoDe 
               Height          =   300
               Left            =   2835
               TabIndex        =   32
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   4215
               TabIndex        =   5
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataEmissaoAte 
               Height          =   300
               Left            =   5385
               TabIndex        =   33
               Top             =   360
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
               TabIndex        =   35
               Top             =   420
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
               TabIndex        =   34
               Top             =   390
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Pedidos de Cotação"
            Height          =   930
            Left            =   570
            TabIndex        =   28
            Top             =   300
            Width           =   6420
            Begin MSMask.MaskEdBox PedidoDe 
               Height          =   300
               Left            =   1725
               TabIndex        =   0
               Top             =   420
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PedidoAte 
               Height          =   300
               Left            =   4200
               TabIndex        =   1
               Top             =   405
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label PedidoDeLabel 
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
               TabIndex        =   30
               Top             =   435
               Width           =   315
            End
            Begin VB.Label PedidoAteLabel 
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
               TabIndex        =   29
               Top             =   465
               Width           =   360
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Validade"
            Height          =   930
            Left            =   570
            TabIndex        =   27
            Top             =   3330
            Width           =   6420
            Begin MSMask.MaskEdBox DataValidadeDe 
               Height          =   300
               Left            =   1620
               TabIndex        =   6
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataValidadeDe 
               Height          =   300
               Left            =   2775
               TabIndex        =   39
               Top             =   390
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataValidadeAte 
               Height          =   300
               Left            =   4215
               TabIndex        =   7
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataValidadeAte 
               Height          =   300
               Left            =   5370
               TabIndex        =   40
               Top             =   390
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               TabIndex        =   42
               Top             =   435
               Width           =   360
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
               TabIndex        =   41
               Top             =   435
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Fornecedor"
            Height          =   930
            Left            =   570
            TabIndex        =   24
            Top             =   1305
            Width           =   6420
            Begin MSMask.MaskEdBox FornecedorDe 
               Height          =   300
               Left            =   1785
               TabIndex        =   2
               Top             =   375
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornecedorAte 
               Height          =   300
               Left            =   4260
               TabIndex        =   3
               Top             =   375
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label FornecedorDeLabel 
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
               TabIndex        =   26
               Top             =   435
               Width           =   315
            End
            Begin VB.Label FornecedorAteLabel 
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   25
               Top             =   390
               Width           =   360
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8745
      Left            =   150
      TabIndex        =   38
      Top             =   360
      Width           =   16515
      _ExtentX        =   29131
      _ExtentY        =   15425
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos de Cotação"
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
Attribute VB_Name = "BaixaPedCotacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjBaixaPedCotacao As ClassBaixaPedCotacao

'Grid de Pedidos de Cotação
Dim objGridPedCotacao As AdmGrid
Dim iGrid_Baixa_Col As Integer
Dim iGrid_PedCotacao_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialForn_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_DataValidade_Col As Integer

'Eventos dos Browses
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorDe As AdmEvento
Attribute objEventoFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorAte As AdmEvento
Attribute objEventoFornecedorAte.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_PedCotacao = 2

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    'Inicializa as variáveis globais
    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoFornecedorDe = New AdmEvento
    Set objEventoFornecedorAte = New AdmEvento
    
    Set objGridPedCotacao = New AdmGrid
    Set gobjBaixaPedCotacao = New ClassBaixaPedCotacao

    'Executa inicializacao do GridPedCotacaos
    lErro = Inicializa_Grid_PedCotacao(objGridPedCotacao)
    If lErro <> SUCESSO Then gError 67550
    
    'Preenche a combo de Ordenação
    Call Ordenados_Carrega
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 67550
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143311)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_PedCotacao(objGridInt As AdmGrid) As Long
'Inicializa o grid de Pedido de Cotação

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Data Validade")

    'campos de edição do grid
    objGridInt.colCampo.Add (Baixa.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataValidade.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Baixa_Col = 1
    iGrid_PedCotacao_Col = 2
    iGrid_Fornecedor_Col = 3
    iGrid_FilialForn_Col = 4
    iGrid_Data_Col = 5
    iGrid_DataEmissao_Col = 6
    iGrid_DataValidade_Col = 7
       
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedCotacao

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDCOTACOES + 1

    'Não permite incluir e excluir novas linhas no grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PedCotacao = SUCESSO

    Exit Function

End Function

Sub Ordenados_Carrega()

    'Limpa a combo de ordenação
    Ordenados.Clear

    'Preenche a combo com as opções
    Ordenados.AddItem "Código"
    Ordenados.AddItem "Fornecedor"
    Ordenados.AddItem "Data de Emissão"
    Ordenados.AddItem "Data de Validade"
                
    Ordenados.ListIndex = 0
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
 
    'Libera as variáveis globais
    Set objEventoPedidoDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoFornecedorDe = Nothing
    Set objEventoFornecedorAte = Nothing

    Set gobjBaixaPedCotacao = Nothing
    Set objGridPedCotacao = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoPedido_Click()
'Chama a tela de Pedido de Cotação

Dim objPedCotacao As New ClassPedidoCotacao

On Error GoTo Erro_BotaoPedido_Click

    'Se nenhuma linha do Grid estiver selecionada, erro
    If GridPedCotacao.Row = 0 Then gError 67629

    'Carrega objPedCotacao com Codigo e FilialEmpresa do Pedido
    objPedCotacao.lCodigo = StrParaLong(GridPedCotacao.TextMatrix(GridPedCotacao.Row, iGrid_PedCotacao_Col))
    objPedCotacao.iFilialEmpresa = giFilialEmpresa
    
    'Chama a tela de Pedido de Compras
    Call Chama_Tela("PedidoCotacao", objPedCotacao)

    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr
        
        Case 67629
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143312)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidadeDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidadeAte_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEmissaoDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 67551

    Exit Sub
                   
Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 67551
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143313)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a DataAte está preenchida
    If Len(Trim(DataEmissaoAte.Text)) = 0 Then Exit Sub

    'Critica a DataAte informada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 67552

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 67552
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143314)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeDe_Validate

    'Verifica se a DataValidadeDe está preenchida
    If Len(Trim(DataValidadeDe.Text)) = 0 Then Exit Sub

    'Critica a DataValidadeDe informada
    lErro = Data_Critica(DataValidadeDe.Text)
    If lErro <> SUCESSO Then gError 67553

    Exit Sub

Erro_DataValidadeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 67553
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143315)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeAte_Validate

    'Verifica se a DataValidadeAte está preenchida
    If Len(Trim(DataValidadeAte.Text)) = 0 Then Exit Sub

    'Critica a DataValidadeAte informada
    lErro = Data_Critica(DataValidadeAte.Text)
    If lErro <> SUCESSO Then gError 67554

    Exit Sub

Erro_DataValidadeAte_Validate:

    Cancel = True

    Select Case gErr

        Case 67554
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143316)

    End Select

    Exit Sub

End Sub

Private Sub PedidoDe_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoDe, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub PedidoAte_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoAte, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub DataEmissaoDe_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub DataEmissaoAte_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub DataValidadeDe_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataValidadeDe, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub DataValidadeAte_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataValidadeAte, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub FornecedorDe_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    iFrameSelecaoAlterado = iFrameAux
    
End Sub

Private Sub FornecedorAte_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFrameSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    iFrameSelecaoAlterado = iFrameAux

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub UpDownDataEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67555

    Exit Sub

Erro_UpDownDataEmissaoDe_DownClick:

    Select Case gErr

        Case 67555
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143317)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67556

    Exit Sub

Erro_UpDownDataEmissaoDe_UpClick:

    Select Case gErr

        Case 67556
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143318)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67557

    Exit Sub

Erro_UpDownDataEmissaoAte_DownClick:

    Select Case gErr

        Case 67557
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143319)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67558

    Exit Sub

Erro_UpDownDataEmissaoAte_UpClick:

    Select Case gErr

        Case 67558
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143320)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidadeDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataValidadeDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67559

    Exit Sub

Erro_UpDownDataValidadeDe_DownClick:

    Select Case gErr

        Case 67559
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143321)

    End Select

    Exit Sub

End Sub
            
Private Sub UpDownDataValidadeDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataValidadeDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67560

    Exit Sub

Erro_UpDownDataValidadeDe_UpClick:

    Select Case gErr

        Case 67560
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143322)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidadeAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataValidadeAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 67561

    Exit Sub

Erro_UpDownDataValidadeAte_DownClick:

    Select Case gErr

        Case 67561
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143323)

    End Select

    Exit Sub

End Sub
            
Private Sub UpDownDataValidadeAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataValidadeAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 67562

    Exit Sub

Erro_UpDownDataValidadeAte_UpClick:

    Select Case gErr

        Case 67562
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143324)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorDe_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoDeLabel_Click()

Dim objPedCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Verifica se PedidoDe esta preenchido
    If Len(Trim(PedidoDe.Text)) > 0 Then

        'Coloca o Codigo do Pedido de Cotação em objPedidoCotacao
        objPedCotacao.lCodigo = StrParaLong(PedidoDe.Text)

    End If

    'Chama a tela PedCotacaoEmitidosLista
    Call Chama_Tela("PedCotacaoEmitidosLista", colSelecao, objPedCotacao, objEventoPedidoDe)

    Exit Sub

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao
    
    Set objPedCotacao = obj1

    'Coloca o codigo retornado em PedidoDe
    PedidoDe.Text = objPedCotacao.lCodigo

    Me.Show

End Sub

Private Sub PedidoAteLabel_Click()

Dim objPedCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

    'Verifica se PedidoAte esta preenchido
    If Len(Trim(PedidoAte.Text)) > 0 Then

        'Coloca o Codigo do Pedido de Cotação em objPedidoCotacao
        objPedCotacao.lCodigo = StrParaLong(PedidoAte.Text)

    End If

    'Chama a tela PedCotacaoEmitidosLista
    Call Chama_Tela("PedCotacaoEmitidosLista", colSelecao, objPedCotacao, objEventoPedidoAte)

    Exit Sub

End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao
    
    Set objPedCotacao = obj1

    'Coloca o codigo retornado em PedidoAte
    PedidoAte.Text = objPedCotacao.lCodigo

    Me.Show

End Sub

Private Sub FornecedorDeLabel_Click()

Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

    'Verifica se FornecedorDe esta preenchido
    If Len(Trim(FornecedorDe.Text)) > 0 Then objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)

     'Chama a tela FornecedorLista
     Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorDe)

End Sub

Private Sub objEventoFornecedorDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o codigo retornado em FornecedorDe
    FornecedorDe.Text = objFornecedor.lCodigo

    Me.Show

End Sub

Private Sub FornecedorAteLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Verifica se FornecedorAte esta preenchido
    If Len(Trim(FornecedorAte.Text)) > 0 Then objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)

     'Chama a tela FornecedorLista
     Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorAte)

End Sub

Private Sub objEventoFornecedorAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o codigo retornado em FornecedorAte
    FornecedorAte.Text = objFornecedor.lCodigo

    Me.Show

End Sub

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim colSaida As New Collection
Dim colCampos As New Collection
Dim colPedMarcados As New Collection
Dim iIndice As Integer
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Ordenados_Click

    'Se o grid não foi preenchido, sai da rotina
    If objGridPedCotacao.iLinhasExistentes = 0 Then Exit Sub
    
    Select Case Ordenados.Text
    
        Case "Código"
            colCampos.Add "lCodigo"
        
        Case "Fornecedor"
            colCampos.Add "lFornecedor"
        
        Case "Data de Emissão"
            colCampos.Add "dtDataEmissao"
        
        Case "Data de Validade"
            colCampos.Add "dtDataValidade"
            
    End Select
        
    'Ordena a coleção
    Call Ordena_Colecao(gobjBaixaPedCotacao.colPedCotacao, colSaida, colCampos)
    Set gobjBaixaPedCotacao.colPedCotacao = colSaida
    
    'Guarda os Pedidos de Cotação marcados
    For iIndice = 1 To objGridPedCotacao.iLinhasExistentes
        If GridPedCotacao.TextMatrix(iIndice, iGrid_Baixa_Col) = "1" Then
            colPedMarcados.Add CLng(GridPedCotacao.TextMatrix(iIndice, iGrid_PedCotacao_Col))
        End If
    Next
    
    Call Grid_Limpa(objGridPedCotacao)
    
    'Preenche o GridPedCotacao
    lErro = Grid_Pedido_Preenche(gobjBaixaPedCotacao.colPedCotacao)
    If lErro <> SUCESSO Then gError 67630
    
    'Marca novamente os Pedidos de Cotação
    For iIndice = 1 To colPedMarcados.Count
        For iLinha = 1 To objGridPedCotacao.iLinhasExistentes
            If CStr(colPedMarcados(iIndice)) = GridPedCotacao.TextMatrix(iLinha, iGrid_PedCotacao_Col) Then
                GridPedCotacao.TextMatrix(iLinha, iGrid_Baixa_Col) = "1"
            End If
        Next
    Next
    
    Call Grid_Refresh_Checkbox(objGridPedCotacao)
    
    Exit Sub
    
Erro_Ordenados_Click:
    
    Select Case gErr
        
        Case 67630
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143325)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedCotacao.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridPedCotacao.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_ATIVO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridPedCotacao)

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedCotacao.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridPedCotacao.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_INATIVO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridPedCotacao)

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
    
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
           
        'Se o frame selecionado foi o de Pedido e houve alteracao do Tab de Selecao
        If TabStrip1.SelectedItem.Index = TAB_PedCotacao And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then
            
            'Recolhe os dados do Tab de Selecao
            lErro = Move_TabSelecao_Memoria()
            If lErro <> SUCESSO Then gError 67563
    
            'Traz para a tela os Pedidos de Cotação com as características determinadas no Tab Selecao
            lErro = Traz_Pedidos_Tela()
            If lErro <> SUCESSO Then gError 67564
                    
            iFrameSelecaoAlterado = 0
            
        End If
 
    End If
    
    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 67563, 67564
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143326)

    End Select

    Exit Sub

End Sub

Private Function Move_TabSelecao_Memoria() As Long
'Recolhe os dados do TabSelecao

On Error GoTo Erro_Move_TabSelecao_Memoria

    'Verifica se PedidoDe e PedidoAte estão preenchidos
    If Len(Trim(PedidoDe.Text)) > 0 And Len(Trim(PedidoAte.Text)) > 0 Then
    
        'Verifica se PedidoDe é maior que PedidoAte
        If (StrParaLong(PedidoDe.Text) > StrParaLong(PedidoAte.Text)) Then gError 67565

    End If
    
    'Recolhe PedidoDe e PedidoAte
    gobjBaixaPedCotacao.lPedCotacaoDe = StrParaLong(PedidoDe.Text)
    gobjBaixaPedCotacao.lPedCotacaoAte = StrParaLong(PedidoAte.Text)

    'Verifica se FornecedorDe e FornecedorAte estão preenchidos
    If Len(Trim(FornecedorDe.Text)) > 0 And Len(Trim(FornecedorAte.Text)) > 0 Then
    
        'Verifica se FornecedorDe é maior que FornecedorAte
        If (StrParaLong(FornecedorDe.Text) > StrParaLong(FornecedorAte.Text)) Then gError 67566

    End If
    
    'Recolhe FornecedorDe e FornecedorAte
    gobjBaixaPedCotacao.lFornecedorDe = StrParaLong(FornecedorDe.Text)
    gobjBaixaPedCotacao.lFornecedorAte = StrParaLong(FornecedorAte.Text)
    
    'Verifica se DataDe e DataAte estão preenchidas
    If Len(Trim(DataEmissaoDe.ClipText)) > 0 And Len(Trim(DataEmissaoAte.ClipText)) > 0 Then
    
        'Verifica se DataEmissaoDe é maior que DataEmissaoAte
        If (StrParaDate(DataEmissaoDe.Text) > StrParaDate(DataEmissaoAte.Text)) Then gError 67567

    End If
    
    'Recolhe DataDe e DataAte
    gobjBaixaPedCotacao.dtDataEmissaoDe = StrParaDate(DataEmissaoDe.Text)
    gobjBaixaPedCotacao.dtDataEmissaoAte = StrParaDate(DataEmissaoAte.Text)

    'Verifica se DataValidadeDe e DataValidadeAte estão preenchidas
    If Len(Trim(DataValidadeDe.ClipText)) > 0 And Len(Trim(DataValidadeAte.ClipText)) > 0 Then
    
        'Verifica se DataValidadeDe é maior que DataValidadeAte
        If (StrParaDate(DataValidadeDe.Text) > StrParaDate(DataValidadeAte.Text)) Then gError 67568
    
    End If
    
    'Recolhe DataValidadeDe e DataValidadeAte
    gobjBaixaPedCotacao.dtDataValidadeDe = StrParaDate(DataValidadeDe.Text)
    gobjBaixaPedCotacao.dtDataValidadeAte = StrParaDate(DataValidadeAte.Text)

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 67565
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)

        Case 67566
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)

        Case 67567
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", gErr)
            
        Case 67568
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143327)

    End Select

    Exit Function

End Function

Function Traz_Pedidos_Tela() As Long
'Traz para a tela os Pedidos de Compra com as características atribuídas no tab Selecao

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Traz_Pedidos_Tela

    'Limpa a colecao de Pedidos
    Set gobjBaixaPedCotacao.colPedCotacao = New Collection

    'Limpa o GridPedCotacao
    Call Grid_Limpa(objGridPedCotacao)

    'Le todos os Pedidos de Cotação com as caracteristicas informadas na Selecao
    lErro = CF("BaixaPedCotacao_ObterPedidos", gobjBaixaPedCotacao)
    If lErro <> SUCESSO Then gError 67569

    'Preenche o GridPedCotacao
    lErro = Grid_Pedido_Preenche(gobjBaixaPedCotacao.colPedCotacao)
    If lErro <> SUCESSO Then gError 67598
    
    Traz_Pedidos_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedidos_Tela:

    Traz_Pedidos_Tela = gErr

    Select Case gErr

        Case 67569, 67598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143328)

    End Select

    Exit Function

End Function

Sub Move_Pedidos_Memoria()

Dim iIndice As Integer
Dim objPedCotacao As ClassPedidoCotacao

    Set gobjBaixaPedCotacao.colPedCotacao = New Collection
    
    'Para cada linha do Grid
    For iIndice = 1 To objGridPedCotacao.iLinhasExistentes
            
        'Se a linha do Grid estiver selecionada
        If GridPedCotacao.TextMatrix(iIndice, iGrid_Baixa_Col) = "1" Then
        
            Set objPedCotacao = New ClassPedidoCotacao
            objPedCotacao.lCodigo = CLng(GridPedCotacao.TextMatrix(iIndice, iGrid_PedCotacao_Col))
            objPedCotacao.iFilialEmpresa = giFilialEmpresa
            
            'Adiciona o Pedido de Cotação na coleção
            gobjBaixaPedCotacao.colPedCotacao.Add objPedCotacao
            
        End If
        
    Next
    
End Sub

Private Function Grid_Pedido_Preenche(colPedCotacao As Collection) As Long
'Preenche o Grid Pedidos com os dados de colPedCotacao

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objPedidoCotacao As ClassPedidoCotacao
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Grid_Pedido_Preenche

    'Percorre toda a Colecao de PedidoCompra
    For Each objPedidoCotacao In colPedCotacao

        iLinha = iLinha + 1

        'Passa para a tela os dados do PedCompra em questão
        GridPedCotacao.TextMatrix(iLinha, iGrid_PedCotacao_Col) = objPedidoCotacao.lCodigo

        If objPedidoCotacao.lFornecedor <> 0 Then
        
            objFornecedor.lCodigo = objPedidoCotacao.lFornecedor
    
            'Lê o Fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 67599
            
            'Se nao encontrou => erro
            If lErro = 12729 Then gError 67600
    
            'Coloca Codigo e NomeReduzido do Fornecedor no GridPedCotacao
            GridPedCotacao.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        
        End If
        
        If objPedidoCotacao.iFilial <> 0 Then
        
            objFilialFornecedor.iCodFilial = objPedidoCotacao.iFilial
            objFilialFornecedor.lCodFornecedor = objPedidoCotacao.lFornecedor
            
            'Lê a FilialFornecedor
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 67601
            
            'Se nao encontrou => erro
            If lErro = 12929 Then gError 67602
    
            'Coloca Codigo e Nome da Filial do Fornecedor no Grid
            GridPedCotacao.TextMatrix(iLinha, iGrid_FilialForn_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        
        End If
        
        'Datas
        If objPedidoCotacao.dtData <> DATA_NULA Then GridPedCotacao.TextMatrix(iLinha, iGrid_Data_Col) = Format(objPedidoCotacao.dtData, "dd/mm/yy")
        If objPedidoCotacao.dtDataEmissao <> DATA_NULA Then GridPedCotacao.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objPedidoCotacao.dtDataEmissao, "dd/mm/yy")
        If objPedidoCotacao.dtDataValidade <> DATA_NULA Then GridPedCotacao.TextMatrix(iLinha, iGrid_DataValidade_Col) = Format(objPedidoCotacao.dtDataValidade, "dd/mm/yy")

    Next

    'Passa para o Obj o número de PedCompra passados pela Coleção
    objGridPedCotacao.iLinhasExistentes = colPedCotacao.Count

    Grid_Pedido_Preenche = SUCESSO

    Exit Function

Erro_Grid_Pedido_Preenche:

    Grid_Pedido_Preenche = gErr

    Select Case gErr

        Case 67599, 67601

        Case 67600
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 67602
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143329)

    End Select

    Exit Function

End Function

Private Sub BotaoBaixa_Click()
'Baixa o Pedido de Compra selecionado no Grid de Pedidos

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_BotaoBaixa_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se existe pelo menos uma linha selecionada no Grid
    For iLinha = 1 To objGridPedCotacao.iLinhasExistentes
        If GridPedCotacao.TextMatrix(iLinha, iGrid_Baixa_Col) = "1" Then
            Exit For
        End If
    Next
    If iLinha > objGridPedCotacao.iLinhasExistentes Then gError 67603
    
    'Move os Pedidos para a memória
    Call Move_Pedidos_Memoria
    
    'Baixa os Pedidos de Cotação selecionados
    lErro = CF("BaixaPedCotacao_Baixar_Pedidos", gobjBaixaPedCotacao.colPedCotacao)
    If lErro <> SUCESSO Then gError 67604
    
    'Traz novamente os Pedidos para a tela
    lErro = Traz_Pedidos_Tela()
    If lErro <> SUCESSO Then gError 67605
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixa_Click:

    Select Case gErr

        Case 67603
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDCOTACAO_SELECIONADOS", gErr)
                    
        Case 67604, 67605
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143330)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Sub Limpa_Tela_BaixaPedCotacao()

    'Limpa o Grid
    Call Grid_Limpa(objGridPedCotacao)
    
    'Limpa Frame Seleção
    PedidoDe.Text = ""
    PedidoAte.Text = ""
    FornecedorDe.Text = ""
    FornecedorAte.Text = ""
    DataEmissaoDe.PromptInclude = False
    DataEmissaoDe.Text = ""
    DataEmissaoDe.PromptInclude = True
    DataValidadeDe.PromptInclude = False
    DataValidadeDe.Text = ""
    DataValidadeDe.PromptInclude = True
    
    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Baixa de Pedidos de Cotação"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "BaixaPedCotacao"
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is PedidoDe Then
            Call PedidoDeLabel_Click
        ElseIf Me.ActiveControl Is PedidoAte Then
            Call PedidoAteLabel_Click
        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call FornecedorDeLabel_Click
        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call FornecedorAteLabel_Click
        End If
    End If
    
End Sub

'Tratamento do Grid
Private Sub GridPedCotacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPedCotacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedCotacao, iAlterado)
    End If

End Sub

Private Sub GridPedCotacao_EnterCell()

    Call Grid_Entrada_Celula(objGridPedCotacao, iAlterado)

End Sub

Private Sub GridPedCotacao_GotFocus()

    Call Grid_Recebe_Foco(objGridPedCotacao)

End Sub

Private Sub GridPedCotacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedCotacao, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedCotacao, iAlterado)
    End If

End Sub

Private Sub GridPedCotacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPedCotacao)
End Sub

Private Sub GridPedCotacao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPedCotacao)

End Sub

Private Sub GridPedCotacao_RowColChange()

    Call Grid_RowColChange(objGridPedCotacao)

End Sub

Private Sub GridPedCotacao_Scroll()

    Call Grid_Scroll(objGridPedCotacao)

End Sub

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

Private Sub PedidoDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidoDeLabel, Source, X, Y)
End Sub

Private Sub PedidoDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidoDeLabel, Button, Shift, X, Y)
End Sub

Private Sub PedidoAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidoAteLabel, Source, X, Y)
End Sub

Private Sub PedidoAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidoAteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub FornecedorDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorDeLabel, Source, X, Y)
End Sub

Private Sub FornecedorDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorDeLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorAteLabel, Source, X, Y)
End Sub

Private Sub FornecedorAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorAteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
