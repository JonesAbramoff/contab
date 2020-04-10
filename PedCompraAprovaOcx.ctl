VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedCompraAprovaOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8310
      Index           =   2
      Left            =   90
      TabIndex        =   16
      Top             =   675
      Visible         =   0   'False
      Width           =   16680
      Begin VB.Frame FrameItens 
         Caption         =   "Itens"
         Height          =   3660
         Left            =   60
         TabIndex        =   31
         Top             =   4605
         Width           =   16575
         Begin VB.TextBox ObsItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3975
            MaxLength       =   255
            TabIndex        =   37
            Top             =   1410
            Width           =   4455
         End
         Begin VB.TextBox DescProd 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1920
            MaxLength       =   255
            TabIndex        =   33
            Top             =   1065
            Width           =   4000
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   705
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   16395
            _ExtentX        =   28919
            _ExtentY        =   1244
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   240
            Left            =   795
            TabIndex        =   32
            Top             =   1050
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UM 
            Height          =   225
            Left            =   4470
            TabIndex        =   34
            Top             =   1065
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Qtde 
            Height          =   225
            Left            =   5085
            TabIndex        =   35
            Top             =   1080
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox DataLim 
            Height          =   225
            Left            =   6285
            TabIndex        =   36
            Top             =   1035
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoUnit 
            Height          =   225
            Left            =   720
            TabIndex        =   38
            Top             =   1455
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   2190
            TabIndex        =   39
            Top             =   1530
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
      End
      Begin VB.Frame Frame7 
         Caption         =   "Pedidos não enviados"
         Height          =   4545
         Left            =   60
         TabIndex        =   23
         Top             =   15
         Width           =   16530
         Begin VB.CommandButton BotaoPedCompras 
            Caption         =   "Pedido de Compras..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7290
            TabIndex        =   7
            Top             =   3900
            Width           =   1830
         End
         Begin VB.CommandButton BotaoDesmarcarTodosReq 
            Caption         =   "Desmarcar Todos"
            Height          =   555
            Left            =   1680
            Picture         =   "PedCompraAprovaOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3900
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodosReq 
            Caption         =   "Marcar Todos"
            Height          =   555
            Left            =   60
            Picture         =   "PedCompraAprovaOcx.ctx":11E2
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3900
            Width           =   1425
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   240
            Left            =   6345
            TabIndex        =   20
            Top             =   1005
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CheckBox Enviar 
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
            Left            =   360
            TabIndex        =   17
            Top             =   930
            Width           =   915
         End
         Begin VB.TextBox ObsPed 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4095
            MaxLength       =   255
            TabIndex        =   21
            Top             =   2550
            Width           =   6900
         End
         Begin MSMask.MaskEdBox Ped 
            Height          =   225
            Left            =   2745
            TabIndex        =   18
            Top             =   975
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   225
            Left            =   4755
            TabIndex        =   19
            Top             =   990
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridPedidos 
            Height          =   705
            Left            =   60
            TabIndex        =   4
            Top             =   210
            Width           =   16380
            _ExtentX        =   28893
            _ExtentY        =   1244
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Filial 
            Height          =   240
            Left            =   750
            TabIndex        =   40
            Top             =   1275
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Comprador 
            Height          =   240
            Left            =   2610
            TabIndex        =   41
            Top             =   1425
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8265
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   16665
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   3960
         Left            =   870
         TabIndex        =   24
         Top             =   390
         Width           =   7665
         Begin VB.Frame Frame9 
            Caption         =   "Número"
            Height          =   1425
            Left            =   4380
            TabIndex        =   28
            Top             =   1170
            Width           =   2385
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   315
               Left            =   780
               TabIndex        =   2
               Top             =   390
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   315
               Left            =   780
               TabIndex        =   3
               Top             =   960
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label Label14 
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
               Left            =   375
               TabIndex        =   30
               Top             =   450
               Width           =   315
            End
            Begin VB.Label Label12 
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
               Left            =   375
               TabIndex        =   29
               Top             =   1020
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data"
            Height          =   1425
            Left            =   990
            TabIndex        =   25
            Top             =   1170
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   12
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   13
               Top             =   870
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   735
               TabIndex        =   0
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
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   720
               TabIndex        =   1
               Top             =   870
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label3 
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
               Left            =   285
               TabIndex        =   27
               Top             =   960
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Left            =   255
               TabIndex        =   26
               Top             =   420
               Width           =   315
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   15225
      ScaleHeight     =   480
      ScaleWidth      =   1575
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1635
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   45
         Picture         =   "PedCompraAprovaOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "PedCompraAprovaOcx.ctx":2356
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "PedCompraAprovaOcx.ctx":2888
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8760
      Left            =   45
      TabIndex        =   14
      Top             =   330
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   15452
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos"
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
Attribute VB_Name = "PedCompraAprovaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer

Dim gobjPedCompraEnvio As ClassPedCompraEnvio

'GridPedidos
Dim objGridPedidos As AdmGrid
Dim iGrid_Enviar_Col As Integer
Dim iGrid_Ped_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialForn_Col As Integer
Dim iGrid_Comprador_Col As Integer
Dim iGrid_ObsPed_Col As Integer

'GridItens
Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProd_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Qtde_Col As Integer
Dim iGrid_PrecoUnit_Col As Integer
Dim iGrid_PrecoTot_Col As Integer
Dim iGrid_DataLim_Col As Integer
Dim iGrid_ObsItem_Col As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Set objGridPedidos = New AdmGrid
    Set objGridItens = New AdmGrid
    Set gobjPedCompraEnvio = New ClassPedCompraEnvio

    'Inicializa o GridPedidos
    lErro = Inicializa_Grid_Pedidos(objGridPedidos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa o GridItens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211135)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objGridPedidos = Nothing
    Set objGridItens = Nothing
    Set gobjPedCompraEnvio = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211136)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Limite")
    objGridInt.colColuna.Add ("Observação")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProd.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Qtde.Name)
    objGridInt.colCampo.Add (PrecoUnit.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (DataLim.Name)
    objGridInt.colCampo.Add (ObsItem.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProd_Col = 2
    iGrid_UM_Col = 3
    iGrid_Qtde_Col = 4
    iGrid_PrecoUnit_Col = 5
    iGrid_PrecoTot_Col = 6
    iGrid_DataLim_Col = 7
    iGrid_ObsItem_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161194)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodosReq_Click()
    Call Grid_Marca_Desmarca(objGridPedidos, iGrid_Enviar_Col, DESMARCADO)
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Sub Limpa_Tela_PedCompra()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PedCompra

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridPedidos)
    Call Grid_Limpa(objGridItens)

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Limpa_Tela_PedCompra:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211137)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela
    Call Limpa_Tela_PedCompra

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 211138)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa o restante da tela
    Call Limpa_Tela_PedCompra

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211139)

    End Select

    Exit Sub

End Sub

Private Sub CodigoAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodigoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)
End Sub

Private Sub CodigoDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodigoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)
End Sub


Private Sub DataAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub

Private Sub DataDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub Enviar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPedidos)
End Sub

Private Sub Enviar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPedidos)
End Sub

Private Sub Enviar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPedidos.objControle = Enviar
    lErro = Grid_Campo_Libera_Foco(objGridPedidos)
    If lErro <> SUCESSO Then Cancel = True

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

    'Se o frame anterior foi o de Seleção e ele foi alterado
    If iFrameAtual <> 1 And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

        'Traz os dados das Pedidos e seus itens para a tela
        lErro = Traz_Pedidos_Tela()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iFrameSelecaoAlterado = 0

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211140)

    End Select

    Exit Sub

End Sub

Private Function Traz_Pedidos_Tela() As Long

Dim lErro As Long, lCodigoPV As Long
Dim objPedCompra As ClassPedidoCompras
Dim iIndice As Integer, iLinha As Integer
Dim objFornecedor As ClassFornecedor
Dim objFilialForn As ClassFilialFornecedor
Dim objComprador As ClassComprador

On Error GoTo Erro_Traz_Pedidos_Tela

    lErro = Move_TabSelecao_Memoria()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("PedComprasEnvio_Le", gobjPedCompraEnvio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gobjPedCompraEnvio.colPedidos.Count = 0 Then gError 211141
    
    Call Grid_Limpa(objGridPedidos)
    
    If gobjPedCompraEnvio.colPedidos.Count >= objGridPedidos.objGrid.Rows Then
        Call Refaz_Grid(objGridPedidos, gobjPedCompraEnvio.colPedidos.Count)
    End If
    
    iLinha = 0
    For Each objPedCompra In gobjPedCompraEnvio.colPedidos

        iLinha = iLinha + 1

        GridPedidos.TextMatrix(iLinha, iGrid_Ped_Col) = CStr(objPedCompra.lCodigo)

        'Verifica se Data é diferente de Data Nula
        If objPedCompra.dtData <> DATA_NULA Then GridPedidos.TextMatrix(iLinha, iGrid_Data_Col) = Format(objPedCompra.dtData, "dd/mm/yyyy")

        Set objFornecedor = New ClassFornecedor
        
        objFornecedor.lCodigo = objPedCompra.lFornecedor

        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM

        GridPedidos.TextMatrix(iLinha, iGrid_Fornecedor_Col) = CStr(objFornecedor.lCodigo) & SEPARADOR & objFornecedor.sNomeReduzido

        Set objFilialForn = New ClassFilialFornecedor

        objFilialForn.iCodFilial = objPedCompra.iFilial
        objFilialForn.lCodFornecedor = objPedCompra.lFornecedor
        
        'Lê a Filial Fornecedor cujo código foi informado
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
            
        GridPedidos.TextMatrix(iLinha, iGrid_FilialForn_Col) = CStr(objFilialForn.iCodFilial) & SEPARADOR & objFilialForn.sNome
            
        If objPedCompra.iComprador <> 0 Then
            Set objComprador = New ClassComprador
            objComprador.iCodigo = objPedCompra.iComprador
            lErro = CF("Comprador_Le", objComprador)
            If lErro <> SUCESSO And lErro <> 50064 Then gError ERRO_SEM_MENSAGEM
            objPedCompra.sComprador = objComprador.sCodUsuario
        End If
            
        GridPedidos.TextMatrix(iLinha, iGrid_Comprador_Col) = objPedCompra.sComprador

        'Preenche a Observacao
        GridPedidos.TextMatrix(iLinha, iGrid_ObsPed_Col) = objPedCompra.sObservacao
                       
    
    Next
    
    objGridPedidos.iLinhasExistentes = gobjPedCompraEnvio.colPedidos.Count
    
    Call Grid_Refresh_Checkbox(objGridPedidos)
    
    Traz_Pedidos_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedidos_Tela:

    Traz_Pedidos_Tela = gErr

    Select Case gErr
    
        Case 211141
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_SEM_DADOS", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211142)

    End Select

    Exit Function

End Function

Private Function Traz_ItensReq_Tela(ByVal iLinha As Integer) As Long

Dim lErro As Long, sProdMask As String
Dim objPedCompra As ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim iIndice As Integer

On Error GoTo Erro_Traz_ItensReq_Tela

    FrameItens.Caption = "Itens"

    If objGridItens.iLinhasExistentes <> 0 Then Call Grid_Limpa(objGridItens)
    
    If Not (gobjPedCompraEnvio Is Nothing) Then

        If iLinha > 0 And iLinha <= gobjPedCompraEnvio.colPedidos.Count Then
        
            Set objPedCompra = gobjPedCompraEnvio.colPedidos.Item(iLinha)
        
            FrameItens.Caption = "Itens - " & CStr(objPedCompra.lCodigo)
        
            iIndice = 0
            For Each objItemPC In objPedCompra.colItens
                iIndice = iIndice + 1
               
                Call Mascara_RetornaProdutoTela(objItemPC.sProduto, sProdMask)
           
                GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdMask
                GridItens.TextMatrix(iIndice, iGrid_DescProd_Col) = objItemPC.sDescProduto
                If objItemPC.dtDataLimite <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataLim_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yyyy")

                GridItens.TextMatrix(iIndice, iGrid_ObsItem_Col) = objItemPC.sObservacao
                GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemPC.sUM
                GridItens.TextMatrix(iIndice, iGrid_Qtde_Col) = Formata_Estoque(objItemPC.dQuantidade)
                GridItens.TextMatrix(iIndice, iGrid_PrecoUnit_Col) = Format(objItemPC.dPrecoUnitario, "STANDARD")
                GridItens.TextMatrix(iIndice, iGrid_PrecoTot_Col) = Format(objItemPC.dQuantidade * objItemPC.dPrecoUnitario - objItemPC.dValorDesconto, "STANDARD")
            
            Next
            
            objGridItens.iLinhasExistentes = objPedCompra.colItens.Count
            
        End If
        
    End If
    
    Traz_ItensReq_Tela = SUCESSO

    Exit Function

Erro_Traz_ItensReq_Tela:

    Traz_ItensReq_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211143)

    End Select

    Exit Function
    
End Function

Function Move_TabSelecao_Memoria() As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    gobjPedCompraEnvio.dtDataDe = StrParaDate(DataDe.Text)
    gobjPedCompraEnvio.dtDataAte = StrParaDate(DataAte.Text)
    gobjPedCompraEnvio.lCodigoDe = StrParaLong(CodigoDe.Text)
    gobjPedCompraEnvio.lCodigoAte = StrParaLong(CodigoAte.Text)
    
    If gobjPedCompraEnvio.dtDataDe <> DATA_NULA And gobjPedCompraEnvio.dtDataAte <> DATA_NULA Then
        If gobjPedCompraEnvio.dtDataDe > gobjPedCompraEnvio.dtDataAte Then gError 211144
    End If
    If gobjPedCompraEnvio.lCodigoDe <> 0 And gobjPedCompraEnvio.lCodigoAte <> 0 Then
        If gobjPedCompraEnvio.lCodigoDe > gobjPedCompraEnvio.lCodigoAte Then gError 211147
    End If

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 211144
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 211147
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211148)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcarTodosReq_Click()
    Call Grid_Marca_Desmarca(objGridPedidos, iGrid_Enviar_Col, MARCADO)
End Sub

Private Sub GridPedidos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPedidos, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If
    
    Exit Sub

End Sub

Private Sub GridPedidos_GotFocus()
    Call Grid_Recebe_Foco(objGridPedidos)
End Sub

Private Sub GridPedidos_EnterCell()
    Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
End Sub

Private Sub GridPedidos_LeaveCell()
    Call Saida_Celula(objGridPedidos)
End Sub

Private Sub GridPedidos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPedidos)
        
End Sub

Private Sub GridPedidos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedidos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedidos, iAlterado)
    End If
    
    Exit Sub
    
End Sub

Private Sub GridPedidos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPedidos)
End Sub

Private Sub GridPedidos_RowColChange()
    Call Grid_RowColChange(objGridPedidos)
    Call Traz_ItensReq_Tela(GridPedidos.Row)
End Sub

Private Sub GridPedidos_Scroll()
    Call Grid_Scroll(objGridPedidos)
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim lErro As Long

On Error GoTo Erro_GridItens_Click

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    Exit Sub

Erro_GridItens_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211149)

    End Select

    Exit Sub

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

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
        
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
        
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 211150

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 211150
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211151)

    End Select

    Exit Function

End Function

Private Sub UpDownDataAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211156)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211157)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211162)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211163)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Pedidos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Pedidos

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Pedidos

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Aprovar")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Comprador")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (Enviar.Name)
    objGridInt.colCampo.Add (Ped.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Comprador.Name)
    objGridInt.colCampo.Add (ObsPed.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Enviar_Col = 1
    iGrid_Ped_Col = 2
    iGrid_Data_Col = 3
    iGrid_Fornecedor_Col = 4
    iGrid_FilialForn_Col = 5
    iGrid_Comprador_Col = 6
    iGrid_ObsPed_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedidos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDIDOS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12
    
    'Largura da primeira coluna
    GridPedidos.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pedidos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Pedidos:

    Inicializa_Grid_Pedidos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211164)

    End Select

    Exit Function

End Function

Private Sub BotaoPedCompras_Click()
'Chama a tela PedComprasEnv

Dim objPC As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedCompras_Click

    'Verifica se existe alguma linha selecionada no GridPedidos
    If GridPedidos.Row = 0 Then gError 211165

    objPC.lCodigo = StrParaLong(GridPedidos.TextMatrix(GridPedidos.Row, iGrid_Ped_Col))
    objPC.iFilialEmpresa = giFilialEmpresa

    'Chama a tela PedComprasEnv
    Call Chama_Tela("PedComprasCons", objPC)

    Exit Sub

Erro_BotaoPedCompras_Click:

    Select Case gErr
    
        Case 211165
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211166)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se  DataDe foi preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211171)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se  DataAte foi preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica DataAte
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211172)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Aprovação de Pedidos de Compras"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PedCompraAprova"

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

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Public Function Gravar_Registro() As Long
'Grava a Concorrencia

Dim lErro As Long
Dim objPedCompra As ClassPedidoCompras
Dim iCount As Integer, iLinha As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    iCount = 0
    iLinha = 0
    For Each objPedCompra In gobjPedCompraEnvio.colPedidos
        iLinha = iLinha + 1
        If StrParaInt(GridPedidos.TextMatrix(iLinha, iGrid_Enviar_Col)) = MARCADO Then
            iCount = iCount + 1
            objPedCompra.iSelecionado = MARCADO
        Else
            objPedCompra.iSelecionado = DESMARCADO
        End If
    Next
    
    If iCount = 0 Then gError 211173
    
    lErro = CF("PedCompraAprova_Grava", gobjPedCompraEnvio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 211173
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_SELECIONADA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211174)

    End Select

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
