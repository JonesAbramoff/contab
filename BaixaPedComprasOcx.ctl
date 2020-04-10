VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPedComprasOcx 
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
      Height          =   8400
      Index           =   2
      Left            =   165
      TabIndex        =   29
      Top             =   495
      Visible         =   0   'False
      Width           =   16485
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   690
         Left            =   2490
         Picture         =   "BaixaPedComprasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   7620
         Width           =   1800
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   690
         Left            =   270
         Picture         =   "BaixaPedComprasOcx.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   7620
         Width           =   1800
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Consultar Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5805
         Picture         =   "BaixaPedComprasOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   7620
         Width           =   1800
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "BaixaPedComprasOcx.ctx":2CCA
         Left            =   1950
         List            =   "BaixaPedComprasOcx.ctx":2CD7
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   300
         Width           =   3480
      End
      Begin VB.TextBox DataEnvio 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6270
         TabIndex        =   40
         Text            =   "Data Envio"
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   3555
         TabIndex        =   34
         Text            =   "Filial"
         Top             =   1395
         Width           =   1065
      End
      Begin VB.TextBox Fornecedor 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   11280
         TabIndex        =   37
         Text            =   "Fornecedor"
         Top             =   3525
         Width           =   2535
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1320
         TabIndex        =   36
         Text            =   "Pedido"
         Top             =   2070
         Width           =   975
      End
      Begin VB.CheckBox Baixa 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   1485
         TabIndex        =   33
         Top             =   1515
         Width           =   870
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
         Height          =   585
         Left            =   6090
         Picture         =   "BaixaPedComprasOcx.ctx":2D07
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   75
         Width           =   1830
      End
      Begin VB.TextBox ValorProdutos 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4395
         TabIndex        =   38
         Text            =   "Valor Produtos Pedido"
         Top             =   2325
         Width           =   1785
      End
      Begin VB.TextBox ValorProdutosRecebido 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5940
         TabIndex        =   39
         Text            =   "Valor Produtos Recebido"
         Top             =   1905
         Width           =   1965
      End
      Begin VB.TextBox Data 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   4860
         TabIndex        =   35
         Text            =   "Data"
         Top             =   1710
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedido 
         Height          =   5880
         Left            =   990
         TabIndex        =   32
         Top             =   825
         Width           =   13245
         _ExtentX        =   23363
         _ExtentY        =   10372
         _Version        =   393216
         Rows            =   11
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox MotivoBaixa 
         Height          =   300
         Left            =   1770
         TabIndex        =   42
         Top             =   7245
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
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
         Left            =   600
         TabIndex        =   30
         Top             =   345
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Motivo da Baixa: "
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
         Left            =   300
         TabIndex        =   41
         Top             =   7275
         Width           =   1500
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
      Left            =   15615
      Picture         =   "BaixaPedComprasOcx.ctx":2E6D
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Fechar"
      Top             =   30
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8430
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   495
      Width           =   16500
      Begin VB.CheckBox ExibeTodos 
         Caption         =   "Exibe Todos os Pedidos"
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
         Left            =   2595
         TabIndex        =   3
         Top             =   600
         Width           =   2430
      End
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Pedidos"
         Height          =   5055
         Left            =   660
         TabIndex        =   2
         Top             =   390
         Width           =   6450
         Begin VB.Frame Frame3 
            Caption         =   "Data"
            Height          =   930
            Left            =   555
            TabIndex        =   15
            Top             =   2715
            Width           =   5505
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   1125
               TabIndex        =   17
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
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   2265
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   3750
               TabIndex        =   20
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
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   4920
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label DataAteLabel 
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
               Left            =   3315
               TabIndex        =   19
               Top             =   420
               Width           =   360
            End
            Begin VB.Label DataDeLabel 
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
               Left            =   690
               TabIndex        =   16
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   1155
            Left            =   555
            TabIndex        =   4
            Top             =   540
            Width           =   5520
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
               Left            =   690
               TabIndex        =   5
               Top             =   300
               Width           =   1905
            End
            Begin MSMask.MaskEdBox PedidoDe 
               Height          =   300
               Left            =   1200
               TabIndex        =   7
               Top             =   690
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
               Left            =   3810
               TabIndex        =   9
               Top             =   690
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
               Left            =   705
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   6
               Top             =   750
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
               Left            =   3330
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   8
               Top             =   705
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Fornecedores"
            Height          =   855
            Left            =   555
            TabIndex        =   10
            Top             =   1755
            Width           =   5520
            Begin MSMask.MaskEdBox FornecedorDe 
               Height          =   300
               Left            =   1200
               TabIndex        =   12
               Top             =   360
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
               Left            =   3810
               TabIndex        =   14
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
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
               Left            =   3345
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   13
               Top             =   375
               Width           =   360
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
               Left            =   675
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   11
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data Envio"
            Height          =   900
            Left            =   555
            TabIndex        =   22
            Top             =   3735
            Width           =   5505
            Begin MSMask.MaskEdBox DataEnvioDe 
               Height          =   300
               Left            =   1080
               TabIndex        =   24
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
            Begin MSComCtl2.UpDown UpDownEnvioDe 
               Height          =   300
               Left            =   2220
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioAte 
               Height          =   300
               Left            =   3735
               TabIndex        =   27
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
            Begin MSComCtl2.UpDown UpDownEnvioAte 
               Height          =   300
               Left            =   4890
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label DataEnvioDeLabel 
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
               Left            =   645
               TabIndex        =   23
               Top             =   420
               Width           =   315
            End
            Begin VB.Label DataEnvioAteLabel 
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
               Left            =   3285
               TabIndex        =   26
               Top             =   420
               Width           =   360
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8880
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   16725
      _ExtentX        =   29501
      _ExtentY        =   15663
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
Attribute VB_Name = "BaixaPedComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorDe As AdmEvento
Attribute objEventoFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorAte As AdmEvento
Attribute objEventoFornecedorAte.VB_VarHelpID = -1

'Grid de Pedidos de Compra
Dim objGridPedido As AdmGrid
Dim iGrid_Baixa_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_DataEnvio_Col As Integer
Dim iGrid_ValorPedido_col As Integer
Dim iGrid_ValorRecebido_Col As Integer

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFramePrincipalAlterado As Integer
Dim iTabSelecaoAlterado As Integer
Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String
Dim gsOrdenacao As String
Dim iBaixaAlterado As Integer
Dim gobjBaixaPedCompras As ClassBaixaPedCompra

Private Sub Baixa_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoBaixa_Click()
'Baixa o Pedido de Compra selecionado no Grid de Pedidos

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim colPedCompras As New Collection
Dim objPedidoCompra As New ClassPedidoCompras
Dim sPedidos As String
Dim vbResult As VbMsgBoxResult
Dim lPedidos(1 To NUM_MAX_PEDIDOS) As Long

On Error GoTo Erro_BotaoBaixa_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedido.iLinhasExistentes

        'Verifica se tem algum pedido marcado
        If GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = MARCADO Then

            Set objPedidoCompra = New ClassPedidoCompras

            'Guarda Código e FilialEmpresa do Pedido de Compra
            objPedidoCompra.lCodigo = StrParaLong(GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col))
            objPedidoCompra.iFilialEmpresa = giFilialEmpresa

            'Guarda o Motivo da Baixa
            objPedidoCompra.sMotivoBaixa = MotivoBaixa.Text
            objPedidoCompra.iTipoBaixa = BAIXA_MANUAL_PEDCOMPRA
            'Adiciona objPedidoCompra na coleção de Pedidos de Compra
            colPedCompras.Add objPedidoCompra

            iIndice = iIndice + 1
            
            lPedidos(iIndice) = objPedidoCompra.lCodigo

        End If

    Next

    'Se não há nenhum pedido marcado ==> erro
    If iIndice = 0 Then gError 56497
    
    sPedidos = ""
    For iLinha = 1 To iIndice
    
        Select Case iLinha
        
            Case 1
                sPedidos = CStr(lPedidos(iLinha))
                
            Case iIndice
                sPedidos = sPedidos & " e " & CStr(lPedidos(iLinha))
            
            Case Else
                sPedidos = sPedidos & ", " & CStr(lPedidos(iLinha))
        
        End Select
    
    Next
    
    vbResult = Rotina_Aviso(vbYesNo, "AVISO_BAIXA_PEDIDOSCOMPRA", sPedidos)
    If vbResult = vbNo Then gError 180233

    'Chama PedidoComprasBaixar_Batch()
    lErro = CF("PedidoComprasBaixar_Batch", colPedCompras)
    If lErro <> SUCESSO Then gError 56498

    'Limpa MotivoBaixa
    MotivoBaixa.Text = ""

    'Traz os Pedidos de Compra não baixados para a tela
    lErro = Traz_Pedidos_Tela()
    If lErro <> SUCESSO Then gError 56499

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixa_Click:

    Select Case gErr

        Case 56497
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDO_BAIXAR", gErr)

        Case 56498, 56499, 180233

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143287)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedido.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_INATIVO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridPedido)

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridPedido.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = GRID_CHECKBOX_ATIVO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridPedido)

    Exit Sub

End Sub

Private Sub BotaoPedido_Click()
'Chama a tela de Pedido de Compras com o Pedido de Compras selecionado

Dim objPedidoCompra As New ClassPedidoCompras
    
On Error GoTo Erro_BotaoPedido_Click
    
    'Verifica se alguma linha do GridPedido esta selecionada
    If GridPedido.Row = 0 Then gError 89436

    'Carrega objPedidoCompra com Codigo e FilialEmpresa do Pedido
    objPedidoCompra.lCodigo = StrParaLong(GridPedido.TextMatrix(GridPedido.Row, iGrid_Pedido_Col))
    objPedidoCompra.iFilialEmpresa = Codigo_Extrai(GridPedido.TextMatrix(GridPedido.Row, iGrid_Filial_Col))
    objPedidoCompra.lNumIntDoc = gobjBaixaPedCompras.colPedCompras.Item(GridPedido.Row).lNumIntDoc

    'Chama a tela  PedComprasCons
    Call Chama_Tela("PedComprasCons", objPedidoCompra)

    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr
    
        Case 89436
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143288)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

Dim iTabSelecao As Integer
    
    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se  DataDe foi preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then Error 56478

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case Err

        Case 56478
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143289)

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
    If lErro <> SUCESSO Then Error 56479

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case Err

        Case 56479
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143290)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEnvioAte_GotFocus()
    
Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataEnvioDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEnvioDe_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se  DataEnvioDe foi preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica DataEnvioDe
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then Error 56480

    Exit Sub

Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case Err

        Case 56480
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143291)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se  DataEnvioAte foi preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica DataEnvioAte
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then Error 56481

    Exit Sub

Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case Err

        Case 56481
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143292)

    End Select

    Exit Sub

End Sub

Private Sub ExibeTodos_Click()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

    'Se marcar ExibeTodos, exibe todos os pedidos
    If ExibeTodos.Value = vbChecked Then
        'Limpa os campos da tela
        PedidoDe.Text = ""
        PedidoAte.Text = ""
        FornecedorDe.Text = ""
        FornecedorAte.Text = ""
        DataDe.PromptInclude = False
        DataDe.Text = ""
        DataDe.PromptInclude = True
        DataAte.PromptInclude = False
        DataAte.Text = ""
        DataAte.PromptInclude = True
        DataEnvioDe.PromptInclude = False
        DataEnvioDe.Text = ""
        DataEnvioDe.PromptInclude = True
        DataEnvioAte.PromptInclude = False
        DataEnvioAte.Text = ""
        DataEnvioAte.PromptInclude = True
        SoResiduais.Value = vbUnchecked
        PedidoDe.Enabled = False
        PedidoAte.Enabled = False
        FornecedorDe.Enabled = False
        FornecedorAte.Enabled = False
        DataDe.Enabled = False
        DataAte.Enabled = False
        DataEnvioDe.Enabled = False
        DataEnvioAte.Enabled = False
        UpDownEmissaoDe.Enabled = False
        UpDownEmissaoAte.Enabled = False
        UpDownEnvioDe.Enabled = False
        UpDownEnvioAte.Enabled = False
        DataDeLabel.Enabled = False
        DataAteLabel.Enabled = False
        PedidoDeLabel.Enabled = False
        PedidoAteLabel.Enabled = False
        FornecedorDeLabel.Enabled = False
        FornecedorAteLabel.Enabled = False
        DataEnvioDeLabel.Enabled = False
        DataEnvioAteLabel.Enabled = False
        SoResiduais.Enabled = False
    Else
        DataDe.Enabled = True
        DataAte.Enabled = True
        PedidoDe.Enabled = True
        PedidoAte.Enabled = True
        SoResiduais.Enabled = True
        DataDeLabel.Enabled = True
        DataEnvioDe.Enabled = True
        DataEnvioAte.Enabled = True
        FornecedorDe.Enabled = True
        DataAteLabel.Enabled = True
        FornecedorAte.Enabled = True
        UpDownEnvioDe.Enabled = True
        PedidoDeLabel.Enabled = True
        PedidoAteLabel.Enabled = True
        UpDownEnvioAte.Enabled = True
        UpDownEmissaoDe.Enabled = True
        DataEnvioDeLabel.Enabled = True
        UpDownEmissaoAte.Enabled = True
        DataEnvioAteLabel.Enabled = True
        FornecedorDeLabel.Enabled = True
        FornecedorAteLabel.Enabled = True
    End If

End Sub

Private Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Private Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    'Preenche a combo de ordenacao com as 3 opcoes abaixo
    asOrdenacao(0) = "PedidoCompra.Codigo, PedidoCompra.NumIntDoc"
    asOrdenacao(1) = "PedidoCompra.Fornecedor, PedidoCompra.FilialDestino, PedidoCompra.Codigo"
    asOrdenacao(2) = "PedidoCompra.Data, PedidoCompra.Codigo"

    asOrdenacaoString(0) = "Código do Pedido"
    asOrdenacaoString(1) = "Fornecedor + FilialFornecedor"
    asOrdenacaoString(2) = "Data do Pedido"

    iFrameAtual = 1

    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoFornecedorDe = New AdmEvento
    Set objEventoFornecedorAte = New AdmEvento
    Set objGridPedido = New AdmGrid
    Set gobjBaixaPedCompras = New ClassBaixaPedCompra

    'Executa inicializacao do GridPedidos
    lErro = Inicializa_Grid_PedCompras(objGridPedido)
    If lErro <> SUCESSO Then Error 56490

    'Limpa a Combobox Ordenados
    Ordenados.Clear

    'Carrega a Combobox Ordenados
    For iIndice = 0 To 2

        Ordenados.AddItem asOrdenacaoString(iIndice)

    Next

    Ordenados.ListIndex = 0
    gsOrdenacao = asOrdenacao(0)
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 56490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143293)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_PedCompras(objGridInt As AdmGrid) As Long
'Inicializa o grid de Pedido de Compras da tela

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data de Envio")
    objGridInt.colColuna.Add ("Valor Produtos Pedido")
    objGridInt.colColuna.Add ("Valor Produtos Recebido")

    'campos de edição do grid
    objGridInt.colCampo.Add (Baixa.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (DataEnvio.Name)
    objGridInt.colCampo.Add (ValorProdutos.Name)
    objGridInt.colCampo.Add (ValorProdutosRecebido.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Baixa_Col = 1
    iGrid_Pedido_Col = 2
    iGrid_Fornecedor_Col = 3
    iGrid_Filial_Col = 4
    iGrid_Data_Col = 5
    iGrid_DataEnvio_Col = 6
    iGrid_ValorPedido_col = 7
    iGrid_ValorRecebido_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPedido

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PEDIDOS + 1

    'Não permite incluir e excluir novas linhas no grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PedCompras = SUCESSO

    Exit Function

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Libera as variaveis globais
    Set objEventoPedidoDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoFornecedorDe = Nothing
    Set objEventoFornecedorAte = Nothing
    Set gobjBaixaPedCompras = Nothing
    Set objGridPedido = Nothing

    'Fecha o comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

End Sub

Private Sub FornecedorAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorAte_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub FornecedorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorAte_Validate

    'Verifica se o fornecedor foi preenchido
    If Len(Trim(FornecedorAte.Text)) = 0 Then Exit Sub

    'Passa o código do fornecedor para objFornecedor
    objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)

    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then Error 57498
    'Se nao encontrou => erro
    If lErro = 12729 Then Error 57499

    Exit Sub

Erro_FornecedorAte_Validate:

    Cancel = True

    Select Case Err

        Case 57498
            'Erro tratado na rotina chamada

        Case 57499
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143294)

    End Select

    Exit Sub


End Sub

Private Sub FornecedorAteLabel_Click()
 
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Verifica se FornecedorAte esta preenchido
    If Len(Trim(FornecedorAte.Text)) > 0 Then objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)

     'Chama a tela FornecedorLista
     Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorAte)

End Sub

Private Sub FornecedorDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FornecedorDe_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub FornecedorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorDe_Validate

    'Verifica se o fornecedor foi preenchido
    If Len(Trim(FornecedorDe.Text)) > 0 Then

        'Passa o código do fornecedor para objFornecedor
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)

        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 57496
        'Se nao encontrou => erro
        If lErro = 12729 Then Error 57497


    End If

    Exit Sub

Erro_FornecedorDe_Validate:

    Cancel = True

    Select Case Err

        Case 57496
            'Erro tratado na rotina chamada

        Case 57497
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143295)

    End Select

    Exit Sub

End Sub

Private Sub GridPedido_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPedido, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedido, iAlterado)
    End If

End Sub

Private Sub GridPedido_GotFocus()
    Call Grid_Recebe_Foco(objGridPedido)
End Sub

Private Sub GridPedido_EnterCell()
    Call Grid_Entrada_Celula(objGridPedido, iAlterado)
End Sub

Private Sub GridPedido_LeaveCell()
    Call Saida_Celula(objGridPedido)
End Sub

Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPedido)
End Sub

Private Sub GridPedido_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPedido, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPedido, iAlterado)
    End If

End Sub

Private Sub GridPedido_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridPedido)
End Sub

Private Sub GridPedido_RowColChange()
    Call Grid_RowColChange(objGridPedido)
End Sub

Private Sub GridPedido_Scroll()
    Call Grid_Scroll(objGridPedido)
End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim objPedidoCompra As ClassPedidoCompras

    Set objPedidoCompra = obj1

    If ExibeTodos.Value = 1 Then ExibeTodos.Value = 0

    'Coloca o codigo retornado em PedidoAte
    PedidoAte.Text = objPedidoCompra.lCodigo

    Me.Show

End Sub
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 57172

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 57172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143296)

    End Select

    Exit Function

End Function

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim objPedidoCompra As ClassPedidoCompras

    Set objPedidoCompra = obj1

    If ExibeTodos.Value = 1 Then ExibeTodos.Value = 0

    'Coloca o codigo retornado em PedidoDe
    PedidoDe.Text = objPedidoCompra.lCodigo

    Me.Show

End Sub

Private Sub objEventoFornecedorDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    If ExibeTodos.Value = 1 Then ExibeTodos.Value = 0

    'Coloca o codigo retornado em FornecedorDe
    FornecedorDe.Text = objFornecedor.lCodigo

    Me.Show

End Sub

Private Sub objEventoFornecedorAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    If ExibeTodos.Value = 1 Then ExibeTodos.Value = 0

    'Coloca o codigo retornado em FornecedorAte
    FornecedorAte.Text = objFornecedor.lCodigo

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

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim colCampos As New Collection
Dim colPedMarcados As New Collection
Dim colSaida As New Collection
Dim iIndice As Integer
Dim iLinha As Integer

On Error GoTo Erro_Ordenados_Click

    'Se o grid não foi preenchido, sai da rotina
    If objGridPedido.iLinhasExistentes = 0 Then Exit Sub
    
    
    Select Case Ordenados.Text
    
        Case "Código do Pedido"
            colCampos.Add "lCodigo"
            colCampos.Add "lNumIntDoc"
            gsOrdenacao = asOrdenacao(0)
            
        Case "Fornecedor + FilialFornecedor"
            colCampos.Add "lFornecedor"
            colCampos.Add "iFilial"
            colCampos.Add "lCodigo"
            gsOrdenacao = asOrdenacao(1)
            
        Case "Data do Pedido"
            colCampos.Add "dtData"
            colCampos.Add "lCodigo"
            gsOrdenacao = asOrdenacao(2)
            
    End Select
         
    'Ordena a coleção
    Call Ordena_Colecao(gobjBaixaPedCompras.colPedCompras, colSaida, colCampos)
    Set gobjBaixaPedCompras.colPedCompras = colSaida
    
    'Guarda os Pedidos de Compra marcados
    For iIndice = 1 To objGridPedido.iLinhasExistentes
        If GridPedido.TextMatrix(iIndice, iGrid_Baixa_Col) = "1" Then
            colPedMarcados.Add CLng(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col))
        End If
    Next
    
    Call Grid_Limpa(objGridPedido)
    
    'Preenche o GridPedido
    lErro = Grid_Pedido_Devolve(gobjBaixaPedCompras.colPedCompras)
    If lErro <> SUCESSO Then Error 57001
    
    'Marca novamente os Pedidos de Compra
    For iIndice = 1 To colPedMarcados.Count
        For iLinha = 1 To objGridPedido.iLinhasExistentes
            If CStr(colPedMarcados(iIndice)) = GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) Then
                GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = "1"
            End If
        Next
    Next
    
    Call Grid_Refresh_Checkbox(objGridPedido)
    
    Exit Sub
    
Erro_Ordenados_Click:

    Select Case Err

        Case 57001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143297)

    End Select

    Exit Sub

End Sub

Private Sub PedidoAte_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoAte_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoAte, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub PedidoAteLabel_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

    'Verifica se PedidoAte foi preenchido
    If Len(Trim(PedidoAte.Text)) > 0 Then

        'Coloca o Codigo em objPedidoCompra
        objPedidoCompra.lCodigo = StrParaLong(PedidoAte.Text)

    End If

    'Chama a tela PedidoComprasLista
    Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoPedidoAte)


End Sub

Private Sub PedidoDe_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoDe_GotFocus()
    
Dim iTabSelecao As Integer
    
    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoDe, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub PedidoDeLabel_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

    'Verifica se PedidoDe esta preenchido
    If Len(Trim(PedidoDe.Text)) > 0 Then

        'Coloca o Codigo do Pedido de Compra em objPedidoCompra
        objPedidoCompra.lCodigo = StrParaLong(PedidoDe.Text)

    End If

    'Chama a tela PedidoComprasLista
     Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoPedidoDe)

    Exit Sub

End Sub

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

    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub
    
    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
    
    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se o frame selecionado foi o de Pedido e houve alteracao do Tab de Selecao
    If TabStrip1.SelectedItem.Index = 2 Then
        'And iTabSelecaoAlterado = 0 Then
        If iFramePrincipalAlterado = REGISTRO_ALTERADO Then
            'Recolhe os dados do Tab de Selecao
            lErro = Move_TabSelecao_Memoria()
            If lErro <> SUCESSO Then Error 56491

            'Traz para a tela os Pedidos de Compra com as características determinadas no Tab Selecao
            lErro = Traz_Pedidos_Tela()
            If lErro <> SUCESSO Then Error 56492
        End If

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case 56491, 56492
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143298)

    End Select

    Exit Sub

End Sub

Private Function Move_TabSelecao_Memoria() As Long
'Recolhe os dados do TabSelecao

On Error GoTo Erro_Move_TabSelecao_Memoria

    'Frame Pedidos
    gobjBaixaPedCompras.iSoResiduais = SoResiduais.Value

    'Verifica se PedidoDe e PedidoAte estão preenchidos
    If Len(Trim(PedidoDe.Text)) > 0 And Len(Trim(PedidoAte.Text)) > 0 Then
    
        'Verifica se PedidoDe é maior que PedidoAte
        If (StrParaLong(PedidoDe.Text) > StrParaLong(PedidoAte.Text)) Then Error 57168

    End If
    
    'Recolhe PedidoDe e PedidoAte
    gobjBaixaPedCompras.lPedCompraDe = StrParaLong(PedidoDe.Text)
    gobjBaixaPedCompras.lPedCompraAte = StrParaLong(PedidoAte.Text)

    'Frame Fornecedores
    'Verifica se FornecedorDe e FornecedorAte estão preenchidos
    If Len(Trim(FornecedorDe.Text)) > 0 And Len(Trim(FornecedorAte.Text)) > 0 Then
    
        'Verifica se FornecedorDe é maior que FornecedorAte
        If (StrParaLong(FornecedorDe.Text) > StrParaLong(FornecedorAte.Text)) Then Error 57169

    End If
    
    'Recolhe FornecedorDe e FornecedorAte
    gobjBaixaPedCompras.lFornecedorDe = StrParaLong(FornecedorDe.Text)
    gobjBaixaPedCompras.lFornecedorAte = StrParaLong(FornecedorAte.Text)

    'Frame Data
    
    'Verifica se DataDe e DataAte estão preenchidas
    If Len(Trim(DataDe.ClipText)) > 0 And Len(Trim(DataAte.ClipText)) > 0 Then
    
        'Verifica se DataDe é maior que DataAte
        If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then Error 57170

    End If
    
    'Recolhe DataDe e DataAte
    gobjBaixaPedCompras.dtDataDe = StrParaDate(DataDe.Text)
    gobjBaixaPedCompras.dtDataAte = StrParaDate(DataAte.Text)

    'Frame DataEnvio
    'Verifica se DataEnvioDe e DataEnvioAte estão preenchidas
    If Len(Trim(DataEnvioDe.ClipText)) > 0 And Len(Trim(DataEnvioAte.ClipText)) > 0 Then
    
        'Verifica se DataEnvioDe é maior que DataEnvioAte
        If StrParaDate(DataEnvioDe.Text) > StrParaDate(DataEnvioAte.Text) Then Error 57171
    
    End If
    
    'Recolhe DataEnvioDe e DataEnvioAte
    gobjBaixaPedCompras.dtDataEnvioDe = StrParaDate(DataEnvioDe.Text)
    gobjBaixaPedCompras.dtDataEnvioAte = StrParaDate(DataEnvioAte.Text)

    'Guarda a ordenação
    gobjBaixaPedCompras.sOrdenacao = asOrdenacao(Ordenados.ListIndex)


    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err

    Select Case Err

        Case 57168
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)

        Case 57169
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", Err)

        Case 57170
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)

        Case 57171
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENVIODE_MAIOR_DATAENVIOATE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143299)

    End Select

    Exit Function

End Function

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 56482

    Exit Sub


Erro_UpDownEmissaoAte_DownClick:

    Select Case Err

        Case 56482
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143300)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 56486

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case Err

        Case 56486
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143301)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 56483

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case Err

        Case 56483
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143302)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()
Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 56487

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case Err

        Case 56487
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143303)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEnvioAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 57490

    Exit Sub

Erro_UpDownEnvioAte_DownClick:

    Select Case Err

        Case 57490
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143304)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEnvioAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 56488

    Exit Sub

Erro_UpDownEnvioAte_UpClick:

    Select Case Err

        Case 56488
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143305)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEnvioDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 56485

    Exit Sub

Erro_UpDownEnvioDe_DownClick:

    Select Case Err

        Case 56485
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143306)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEnvioDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 56489

    Exit Sub

Erro_UpDownEnvioDe_UpClick:

    Select Case Err

        Case 56489
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143307)

    End Select

    Exit Sub

End Sub
Function Traz_Pedidos_Tela() As Long
'Traz para a tela os Pedidos de Compra com as características atribuídas no tab Selecao

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Traz_Pedidos_Tela

    'Limpa a colecao de Pedidos
    Set gobjBaixaPedCompras.colPedCompras = New Collection

    'Limpa o GridPedido
    Call Grid_Limpa(objGridPedido)

    'Le todos os Pedidos com as caracteristicas informadas na Selecao
    lErro = CF("BaixaPedCompras_ObterPedidos", gobjBaixaPedCompras)
    If lErro <> SUCESSO Then Error 57000

    'Preenche o GridPedido
    Call Grid_Pedido_Preenche(gobjBaixaPedCompras.colPedCompras)

    'Selecionar todos os pedidos da tela
    For iLinha = 1 To objGridPedido.iLinhasExistentes
        'Marca o Pedido na tela
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = MARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridPedido)

    iFramePrincipalAlterado = 0

    Traz_Pedidos_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedidos_Tela:

    Traz_Pedidos_Tela = Err

    Select Case Err

        Case 57000
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143308)

    End Select

    Exit Function

End Function

Private Function Grid_Pedido_Preenche(colPedCompra As Collection) As Long
'Preenche o Grid Pedido com os dados de colPedCompra

Dim lErro As Long
Dim iLinha As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sNomeRed As String
Dim dValorProdutosRecebidos As Double

On Error GoTo Erro_Grid_Pedido_Preenche

    'Verifica se o número de pedidos encontrados é superior ao máximo permitido
    'If colPedCompra.Count + 1 > NUM_MAX_PEDIDOS Then Error 57487

    'Se o número de PedCompra for maior que o número de linhas do Grid
    If colPedCompra.Count + 1 > GridPedido.Rows Then

        'Altera o número de linhas do Grid de acordo com o número de PedCompras
        GridPedido.Rows = colPedCompra.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridPedido)

    End If

    iLinha = 0

    'Percorre toda a Colecao de PedidoCompra
    For Each objPedidoCompra In colPedCompra

        iLinha = iLinha + 1

        'Passa para a tela os dados do PedCompra em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = objPedidoCompra.iTipoBaixa
        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objPedidoCompra.lCodigo

        objFornecedor.lCodigo = objPedidoCompra.lFornecedor

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 57136
        'Se nao encontrou => erro
        If lErro = 12729 Then Error 57137

        'Coloca Codigo e NomeReduzido do Fornecedor no GridPedido
        GridPedido.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objPedidoCompra.lFornecedor & SEPARADOR & objFornecedor.sNomeReduzido

        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial
        sNomeRed = objFornecedor.sNomeReduzido
        'Le a FilialFornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 57138
        'Se nao encontrou => erro
        If lErro = 18272 Then Error 57139

        'Coloca Codigo e Nome da Filial do Fornecedor no Grid
        GridPedido.TextMatrix(iLinha, iGrid_Filial_Col) = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

        If objPedidoCompra.dtDataEnvio <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_DataEnvio_Col) = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yy")
        If objPedidoCompra.dtData <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_Data_Col) = Format(objPedidoCompra.dtData, "dd/mm/yy")

        'Preenche a linha do grid com o valor dos produtos
        GridPedido.TextMatrix(iLinha, iGrid_ValorPedido_col) = Format(objPedidoCompra.dValorProdutos, "Standard")

        'Calcula o valor dos produtos recebidos do Pedido de Compra em questao
        lErro = CF("Valor_Produtos_Recebido", objPedidoCompra, dValorProdutosRecebidos)
        If lErro <> SUCESSO Then Error 57492

        'Preenche a linha do grid com o valor recebido dos produtos obtido
        GridPedido.TextMatrix(iLinha, iGrid_ValorRecebido_Col) = Format(dValorProdutosRecebidos, "Standard")

    Next

    'Passa para o Obj o número de PedCompra passados pela Coleção
    objGridPedido.iLinhasExistentes = colPedCompra.Count

    Grid_Pedido_Preenche = SUCESSO

    Exit Function

Erro_Grid_Pedido_Preenche:

    Grid_Pedido_Preenche = Err

    Select Case Err

        Case 57136, 57138, 57492

        Case 57137
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES1", Err, objPedidoCompra.lFornecedor, objFilialFornecedor.iCodFilial)

        Case 57139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES1", Err, objPedidoCompra.lFornecedor)

        Case 57487
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_PEDIDOS_SELECIONADOS_SUPERIOR_MAXIMO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143309)

    End Select

    Exit Function

End Function

Private Function Grid_Pedido_Devolve(colPedCompra As Collection) As Long
'Preenche o Grid Pedido com os dados de colPedCompra

Dim lErro As Long
Dim iLinha As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objItemPC As New ClassItemPedCompra
Dim sNomeRed As String
Dim dValorProdutosRecebidos As Double
Dim dValorTotal  As Double
Dim dQuantRecebidaEfetiva  As Double
Dim dValorDescontoProporcional  As Double
Dim dValorRecebidoEfetivo  As Double
Dim dValorProdutoRecebido As Double

On Error GoTo Erro_Grid_Pedido_Devolve

    'Verifica se o número de pedidos encontrados é superior ao máximo permitido
    If colPedCompra.Count + 1 > NUM_MAX_PEDIDOS Then gError 68397

    'Se o número de PedCompra for maior que o número de linhas do Grid
    If colPedCompra.Count + 1 > GridPedido.Rows Then

        'Altera o número de linhas do Grid de acordo com o número de PedCompras
        GridPedido.Rows = colPedCompra.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridPedido)

    End If

    iLinha = 0

    'Percorre toda a Colecao de PedidoCompra
    For Each objPedidoCompra In colPedCompra

        iLinha = iLinha + 1

        'Passa para a tela os dados do PedCompra em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = objPedidoCompra.iTipoBaixa
        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objPedidoCompra.lCodigo

        objFornecedor.lCodigo = objPedidoCompra.lFornecedor

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 68398
        'Se nao encontrou => erro
        If lErro = 12729 Then gError 68399

        'Coloca Codigo e NomeReduzido do Fornecedor no GridPedido
        GridPedido.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objPedidoCompra.lFornecedor & SEPARADOR & objFornecedor.sNomeReduzido

        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial
        sNomeRed = objFornecedor.sNomeReduzido
        'Le a FilialFornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 68400
        'Se nao encontrou => erro
        If lErro = 18272 Then gError 68401

        'Coloca Codigo e Nome da Filial do Fornecedor no Grid
        GridPedido.TextMatrix(iLinha, iGrid_Filial_Col) = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

        If objPedidoCompra.dtDataEnvio <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_DataEnvio_Col) = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yy")
        If objPedidoCompra.dtData <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_Data_Col) = Format(objPedidoCompra.dtData, "dd/mm/yy")

        'Preenche a linha do grid com o valor dos produtos
        GridPedido.TextMatrix(iLinha, iGrid_ValorPedido_col) = Format(objPedidoCompra.dValorProdutos, "Standard")

        
        'Lê os itens do Pedido de Compra, cujo número do Pedido foi passado como parâmetro
        For Each objItemPC In objPedidoCompra.colItens

            'Calcula o valor total do Item do Pedido de Compra
            dValorTotal = (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) - objItemPC.dValorDesconto

            'Calcula a quantidade efetiva recebida do item do Pedido de Compra
            dQuantRecebidaEfetiva = objItemPC.dQuantRecebida + objItemPC.dQuantRecebimento

            'Calcula o valor do desconto, proporcional à quantidade já recebida do item do Pedido de Compra,
            'O valor do desconto armazenado no banco de dados está referenciando o desconto total do produto,
            'não considerando a quantidade efetiva recebida
            If objItemPC.dQuantidade <> 0 Then dValorDescontoProporcional = (objItemPC.dValorDesconto * dQuantRecebidaEfetiva) / objItemPC.dQuantidade

            'Calcula o Valor Recebido Efetivo do item do Pedido de Compra
            dValorRecebidoEfetivo = (objItemPC.dPrecoUnitario * dQuantRecebidaEfetiva) - dValorDescontoProporcional

            'Calcula o Valor Recebido dos produtos do Pedido de Compra em questão
            dValorProdutoRecebido = dValorRecebidoEfetivo + dValorProdutoRecebido

        Next
    
        'Preenche a linha do grid com o valor recebido dos produtos obtido
        GridPedido.TextMatrix(iLinha, iGrid_ValorRecebido_Col) = Format(dValorProdutoRecebido, "Standard")
        dValorProdutoRecebido = 0
        
    Next

    'Passa para o Obj o número de PedCompra passados pela Coleção
    objGridPedido.iLinhasExistentes = colPedCompra.Count

    Grid_Pedido_Devolve = SUCESSO

    Exit Function

Erro_Grid_Pedido_Devolve:

    Grid_Pedido_Devolve = gErr

    Select Case gErr

        Case 68398, 68400

        Case 68399
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES1", gErr, objPedidoCompra.lFornecedor, objFilialFornecedor.iCodFilial)

        Case 68401
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES1", gErr, objPedidoCompra.lFornecedor)

        Case 68397
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_PEDIDOS_SELECIONADOS_SUPERIOR_MAXIMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143310)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Baixa de Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BaixaPedCompras"

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

Private Sub DataDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataDeLabel, Source, X, Y)
End Sub

Private Sub DataDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataDeLabel, Button, Shift, X, Y)
End Sub

Private Sub DataAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataAteLabel, Source, X, Y)
End Sub

Private Sub DataAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataAteLabel, Button, Shift, X, Y)
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

Private Sub FornecedorAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorAteLabel, Source, X, Y)
End Sub

Private Sub FornecedorAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorAteLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorDeLabel, Source, X, Y)
End Sub

Private Sub FornecedorDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorDeLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEnvioDeLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEnvioDeLabel, Source, X, Y)
End Sub

Private Sub DataEnvioDeLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEnvioDeLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEnvioAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEnvioAteLabel, Source, X, Y)
End Sub

Private Sub DataEnvioAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEnvioAteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function
