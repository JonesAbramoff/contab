VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPedidoSRVOcx 
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   5475
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   2
      Left            =   150
      TabIndex        =   0
      Top             =   645
      Visible         =   0   'False
      Width           =   8925
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todas"
         Height          =   675
         Left            =   2775
         Picture         =   "BaixaPedidoSRVOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3735
         Width           =   1830
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todas"
         Height          =   675
         Left            =   840
         Picture         =   "BaixaPedidoSRVOcx.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3735
         Width           =   1830
      End
      Begin VB.TextBox NomeRed 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3270
         TabIndex        =   10
         Text            =   "Nome"
         Top             =   1215
         Width           =   1995
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Editar Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6405
         Picture         =   "BaixaPedidoSRVOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3735
         Width           =   1830
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "BaixaPedidoSRVOcx.ctx":2E7A
         Left            =   1545
         List            =   "BaixaPedidoSRVOcx.ctx":2E87
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   263
         Width           =   4470
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   6315
         TabIndex        =   7
         Text            =   "Emissão"
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   5400
         TabIndex        =   6
         Text            =   "Filial"
         Top             =   1230
         Width           =   1290
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2040
         TabIndex        =   5
         Text            =   "Cliente"
         Top             =   1200
         Width           =   795
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1080
         TabIndex        =   4
         Text            =   "Pedido"
         Top             =   1215
         Width           =   750
      End
      Begin VB.CheckBox Baixa 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   270
         TabIndex        =   3
         Top             =   1365
         Width           =   816
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
         Height          =   570
         Left            =   6435
         Picture         =   "BaixaPedidoSRVOcx.ctx":2EB7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   165
         Width           =   1830
      End
      Begin VB.TextBox DataEntrega 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7470
         TabIndex        =   1
         Text            =   "Entrega"
         Top             =   1155
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedido 
         Height          =   2520
         Left            =   75
         TabIndex        =   13
         Top             =   870
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   4445
         _Version        =   393216
         Rows            =   10
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
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
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4620
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   645
      Width           =   8835
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Pedidos"
         Height          =   4500
         Left            =   1110
         TabIndex        =   17
         Top             =   60
         Width           =   6270
         Begin VB.Frame Frame3 
            Caption         =   "Data Emissão"
            Height          =   825
            Left            =   435
            TabIndex        =   36
            Top             =   2490
            Width           =   5520
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   780
               TabIndex        =   37
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
               Left            =   1935
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   390
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   39
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
               Left            =   4590
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Left            =   2985
               TabIndex        =   42
               Top             =   420
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
               Left            =   345
               TabIndex        =   41
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   810
            Left            =   435
            TabIndex        =   31
            Top             =   585
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   810
               TabIndex        =   32
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
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   3450
               TabIndex        =   33
               Top             =   367
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelPedidoDe 
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
               Left            =   330
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   35
               Top             =   405
               Width           =   315
            End
            Begin VB.Label LabelPedidoAte 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   34
               Top             =   420
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Clientes"
            Height          =   825
            Left            =   435
            TabIndex        =   26
            Top             =   1530
            Width           =   5520
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   840
               TabIndex        =   27
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
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   28
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
            Begin VB.Label LabelClienteAte 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   30
               Top             =   375
               Width           =   360
            End
            Begin VB.Label LabelClienteDe 
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
               Left            =   315
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   29
               Top             =   420
               Width           =   315
            End
         End
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
            Left            =   1530
            TabIndex        =   25
            Top             =   270
            Width           =   2430
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data Entrega"
            Height          =   825
            Left            =   435
            TabIndex        =   18
            Top             =   3450
            Width           =   5520
            Begin MSMask.MaskEdBox DataEntregaDe 
               Height          =   300
               Left            =   780
               TabIndex        =   19
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
            Begin MSComCtl2.UpDown UpDownEntregaDe 
               Height          =   300
               Left            =   1950
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEntregaAte 
               Height          =   300
               Left            =   3420
               TabIndex        =   21
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
            Begin MSComCtl2.UpDown UpDownEntregaAte 
               Height          =   300
               Left            =   4590
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label7 
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
               TabIndex        =   24
               Top             =   420
               Width           =   315
            End
            Begin VB.Label Label8 
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
               Left            =   2985
               TabIndex        =   23
               Top             =   420
               Width           =   360
            End
         End
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
      Left            =   7470
      Picture         =   "BaixaPedidoSRVOcx.ctx":301D
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Fechar"
      Top             =   75
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5070
      Left            =   105
      TabIndex        =   43
      Top             =   300
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8943
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
Attribute VB_Name = "BaixaPedidoSRVOcx"
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
Dim iTabPrincipalAlterado As Integer
Dim iClienteAlterado  As Integer
Dim gobjBaixaPedido As New ClassBaixaPedidos

Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String

Dim objGrid As AdmGrid
Dim iGrid_Baixa_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_NomeRed_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Emissao_Col As Integer
Dim iGrid_Entrega_Col As Integer

'Eventos de Browse
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_Pedidos = 2

Private Sub Baixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Baixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Baixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Baixa
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub BotaoBaixa_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim objPVInfo As ClassPVInfo
Dim iUmaLinha_Marcada As Integer

On Error GoTo Erro_BotaoBaixa_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Percorre as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        'Varre a coleção de Pedidos
        For Each objPVInfo In gobjBaixaPedido.colPVInfo
            
            'Se for o Pedido que esta em questao atualiza ele na colecao
            If objPVInfo.lCodPedido = CLng(GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col)) Then

                objPVInfo.iMarcada = CInt(GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col))
                
                Exit For
                                
            End If
            
        Next
        
        'Verifica se tem algum pedido marcado
        If GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = S_MARCADO Then
            iUmaLinha_Marcada = MARCADO
        End If
        
    Next
    
    'Se não há nenhum pedido marcado ==> erro
    If iUmaLinha_Marcada <> MARCADO Then gError 210687
    
    lErro = CF("PedidosSRVBaixar_Batch", gobjBaixaPedido.colPVInfo)
    If lErro <> SUCESSO Then gError 210688

    'Descarrega o Grid com os Pedidos que foram Baixados
    lErro = Descarrega_Grid()
    If lErro <> SUCESSO Then gError 210689

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixa_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 210687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDO_BAIXAR", gErr)

        Case 210688, 210689

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210690)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = S_DESMARCADO

        'Desmarca no Obj o pedido em questão
        gobjBaixaPedido.colPVInfo.Item(iLinha).iMarcada = DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGrid)

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = S_MARCADO

        gobjBaixaPedido.colPVInfo.Item(iLinha).iMarcada = MARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGrid)

End Sub

Private Sub BotaoPedido_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim objPVInfo As New ClassPVInfo
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoPedido_Click
        
    'Tem que selecionar alguma linha
    If GridPedido.Row = 0 Then gError 210691
    
    'Tem que ter pelo menos um pedido No Grid
    If GridPedido.Row > gobjBaixaPedido.colPVInfo.Count Then Exit Sub
    
    'Passa os dados do Grid para o Obj
    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
    objPedidoDeVenda.lCodigo = CLng(GridPedido.TextMatrix(GridPedido.Row, iGrid_Pedido_Col))

    'Chama a tela de Pedidos de Servico
    Call Chama_Tela("PedidoServico", objPedidoDeVenda)

    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr

        Case 210691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210692)

    End Select

    Exit Sub

End Sub

Private Sub ClienteAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ClienteAte_GotFocus()
Dim iTabAux As Integer
Dim iClienteAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    iClienteAux = iClienteAlterado
    
    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
    iTabPrincipalAlterado = iTabAux
    iClienteAlterado = iClienteAux

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)
'Verifica se o Cliente De é maior que o Cliente Até
'Verifica a integridade do cliente com o BD

Dim lErro As Long
Dim objClienteAte As New ClassCliente
Dim iCodFilial As Integer
Dim iCria As Integer
Dim colCodigoNome As AdmColCodigoNome

On Error GoTo Erro_ClienteAte_Validate
    
    If iClienteAlterado = 1 Then

        If Len(Trim(ClienteAte.Text)) > 0 Then
            
            'Se o Cliente De estiver preenchido
            If Len(Trim(ClienteDe.Text)) > 0 Then
                'Verifica se o Cliente De é maior que o Cliente Até ----->>> Erro
                If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then gError 210693
                
            End If
            
            objClienteAte.lCodigo = ClienteAte.Text
            
            'Le o Cliente para testar sua integridade com o BD
            lErro = CF("Cliente_Le", objClienteAte)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 210694
            
            'Se não encontrou ----> erro
            If lErro = 12293 Then gError 210695
            
        End If

        iClienteAlterado = 0

    End If
    
    Exit Sub
    
Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr
    
    Case 210693
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", gErr)
            
    Case 210694
        
    Case 210695
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objClienteAte.lCodigo)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210696)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteDe_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ClienteDe_GotFocus()
Dim iTabAux As Integer
Dim iClienteAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    iClienteAux = iClienteAlterado
    
    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
    iTabPrincipalAlterado = iTabAux
    iClienteAlterado = iClienteAux

End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)
'Verifica se o Cliente De é maior que o Cliente Até
'Verifica a integridade do cliente com o BD

Dim lErro As Long
Dim objClienteDe As New ClassCliente
Dim iCodFilial As Integer
Dim iCria As Integer
Dim colCodigoNome As AdmColCodigoNome

On Error GoTo Erro_ClienteDe_Validate
    
    'Se o Cliente foi alterado
    If iClienteAlterado = 1 Then
            
        'Se o ClienteDe estiver Preenchido
        If Len(Trim(ClienteDe.Text)) > 0 Then
            
            'Se o ClienteAte estiver Preenchido
            If Len(Trim(ClienteAte.Text)) > 0 Then
                'Verifica se o CLienteDe é Menor que o ClienteAte
                If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then gError 210697
            End If
            
            objClienteDe.lCodigo = CLng(ClienteDe.Text)
            
            'Lê o Cliente no BD
            lErro = CF("Cliente_Le", objClienteDe)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 210698
            
            'Se não encontrou ---> ERRO
            If lErro = 12293 Then gError 210699
            
        End If

        iClienteAlterado = 0

    End If
    
    Exit Sub
    
Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr
    
    Case 210697
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", gErr)
        
    Case 210698
    
    Case 210699
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objClienteDe.lCodigo)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210700)

    End Select
    
    Exit Sub
    
End Sub

Private Sub DataEmissaoAte_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Se a DataEmissaoAte está preenchida
    If Len(DataEmissaoAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoAte é válida
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 210701

    If Len(Trim(DataEmissaoDe.ClipText)) = 0 Then Exit Sub
    
    'Verifica se a DataEmissaoDe é menor que a DataEmissaoAte
    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then gError 210702

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 210701

        Case 210702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210703)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoDe_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoDe_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 210704

    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é menor que a DataEmissaoAte
    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then gError 210705

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 210704
            
        Case 210705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210706)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaAte_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaAte_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaAte_Validate

    'Se a DataEntregaAte está preenchida
    If Len(DataEntregaAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEntregaAte é válida
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then gError 210707

    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub

    'Verifica se a DataEntregaDe é menor que a DataEntregaAte
    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then gError 210708

    Exit Sub

Erro_DataEntregaAte_Validate:
    
    Cancel = True

    Select Case gErr

        Case 210707
    
        Case 210708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210709)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaDe_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntregaDe_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaDe_Validate

    'Se a DataEntregaDe está preenchida
    If Len(DataEntregaDe.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEntregaDe é válida
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then gError 210710

    If Len(Trim(DataEntregaAte.ClipText)) = 0 Then Exit Sub
    
    'Verifica se a DataEntregaDe é menor que a DataEntregaAte
    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then gError 210711

    Exit Sub

Erro_DataEntregaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 210710

        Case 210711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210712)

    End Select

    Exit Sub

End Sub

Private Sub ExibeTodos_Click()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

    'Limpa os campos da tela
    PedidoInicial.Text = ""
    PedidoFinal.Text = ""
    ClienteDe.Text = ""
    ClienteAte.Text = ""
    DataEmissaoDe.PromptInclude = False
    DataEmissaoDe.Text = ""
    DataEmissaoDe.PromptInclude = True
    DataEmissaoAte.PromptInclude = False
    DataEmissaoAte.Text = ""
    DataEmissaoAte.PromptInclude = True
    DataEntregaDe.PromptInclude = False
    DataEntregaDe.Text = ""
    DataEntregaDe.PromptInclude = True
    DataEntregaAte.PromptInclude = False
    DataEntregaAte.Text = ""
    DataEntregaAte.PromptInclude = True

    'Se marcar ExibeTodos, exibe todos os pedidos
    If ExibeTodos.Value = 1 Then
        PedidoInicial.Enabled = False
        PedidoFinal.Enabled = False
        ClienteDe.Enabled = False
        ClienteAte.Enabled = False
        DataEmissaoDe.Enabled = False
        DataEmissaoAte.Enabled = False
        DataEntregaDe.Enabled = False
        DataEntregaAte.Enabled = False
        UpDownEmissaoDe.Enabled = False
        UpDownEmissaoAte.Enabled = False
        UpDownEntregaDe.Enabled = False
        UpDownEntregaAte.Enabled = False
    Else
        PedidoInicial.Enabled = True
        PedidoFinal.Enabled = True
        ClienteDe.Enabled = True
        ClienteAte.Enabled = True
        DataEmissaoDe.Enabled = True
        DataEmissaoAte.Enabled = True
        DataEntregaDe.Enabled = True
        DataEntregaAte.Enabled = True
        UpDownEmissaoDe.Enabled = True
        UpDownEmissaoAte.Enabled = True
        UpDownEntregaDe.Enabled = True
        UpDownEntregaAte.Enabled = True
    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    'Preenche o Vetor de Ordenação
    asOrdenacao(0) = "PedidoServico.Codigo"
    asOrdenacao(1) = "MIN(PedidoServico.Cliente), PedidoServico.Codigo"
    asOrdenacao(2) = "MIN(PedidoServico.DataEmissao), PedidoServico.Codigo"

    asOrdenacaoString(0) = "Pedido"
    asOrdenacaoString(1) = "Cliente + Pedido"
    asOrdenacaoString(2) = "Data de Emissão do Pedido + Pedido"
    
    'Configura o Frame Atual
    iFrameAtual = 1

    Set objGrid = New AdmGrid

    'Executa a Inicialização do grid Pedido
    lErro = Inicializa_Grid_Pedido(objGrid)
    If lErro <> SUCESSO Then gError 210713

    'Limpa a Combobox Ordenados
    Ordenados.Clear

    'Carrega a Combobox Ordenados
    For iIndice = 0 To 2

        Ordenados.AddItem asOrdenacaoString(iIndice)

    Next
    
    'Configura ordenação
    Ordenados.ListIndex = 0
    
    'Inicializa os Eventos de Browser
    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 210713

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210714)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Pedido(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Pedidos

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixa")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Entrega")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Baixa.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (NomeRed.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)

    'Colunas do Grid
    iGrid_Baixa_Col = 1
    iGrid_Pedido_Col = 2
    iGrid_Cliente_Col = 3
    iGrid_NomeRed_Col = 4
    iGrid_Filial_Col = 5
    iGrid_Emissao_Col = 6
    iGrid_Entrega_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridPedido

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'Largura da primeira coluna
    GridPedido.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'GridPedido.Width = 8400

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Pedido = SUCESSO

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 210715

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 210715
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210716)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoPedidoDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
    
    Set objGrid = Nothing

    Set gobjBaixaPedido = Nothing
    
End Sub

Private Sub GridPedido_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridPedido_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridPedido_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridPedido_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridPedido_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridPedido_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridPedido_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridPedido_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub LabelClienteAte_Click()

Dim colSelecao As Collection
Dim objcliente As New ClassCliente

    'Preenche ClienteAte com o cliente da tela
    If Len(Trim(ClienteAte.Text)) > 0 Then objcliente.lCodigo = CLng(ClienteAte.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteAte)

End Sub

Private Sub LabelClienteDe_Click()

Dim colSelecao As Collection
Dim objcliente As New ClassCliente

    'Preenche ClienteDe com o cliente da tela
    If Len(Trim(ClienteDe.Text)) > 0 Then objcliente.lCodigo = CLng(ClienteDe.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteDe)

End Sub

Private Sub LabelPedidoAte_Click()

Dim colSelecao As Collection
Dim objPedidoDeVenda As New ClassPedidoDeVenda

    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(PedidoFinal.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoFinal.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedSRVAbertosLista
    Call Chama_Tela("PedSRVAbertosLista", colSelecao, objPedidoDeVenda, objEventoPedidoAte)

End Sub

Private Sub LabelPedidoDe_Click()

Dim colSelecao As Collection
Dim objPedidoDeVenda As New ClassPedidoDeVenda

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(PedidoInicial.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoInicial.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedSRVAbertosLista
    Call Chama_Tela("PedSRVAbertosLista", colSelecao, objPedidoDeVenda, objEventoPedidoDe)

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objcliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoClienteAte_evSelecao

    Set objcliente = obj1
    
    If ExibeTodos.Value = 1 Then ExibeTodos.Value = 0
    
    ClienteAte.Text = CStr(objcliente.lCodigo)

    'Chama o Validate de ClienteAte
    Call ClienteAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoClienteAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210717)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objcliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoClienteDe_evSelecao

    Set objcliente = obj1
    
    If ExibeTodos.Value = 1 Then ExibeTodos = 0

    ClienteDe.Text = CStr(objcliente.lCodigo)
    
    'Chama o Validate do Cliente
    Call ClienteDe_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210718)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim bCancel As Boolean

On Error GoTo Erro_objEventoPedidoAte_evSelecao

    Set objPedidoDeVenda = obj1
    
    If ExibeTodos.Value = 1 Then ExibeTodos = 0

    PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)

    'Chama o Validate de PedidoFinal
    Call PedidoFinal_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoPedidoAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210719)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim bCancel As Boolean

On Error GoTo Erro_objEventoPedidoDe_evSelecao

    Set objPedidoDeVenda = obj1
    
    If ExibeTodos.Value = 1 Then ExibeTodos = 0

    PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)
    
    'Chama o validate do PedidoInicial
    Call PedidoInicial_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoPedidoDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210720)

    End Select

    Exit Sub

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex = -1 Then Exit Sub

    'Verifica se a coleção de NFiscal está vazia
    If gobjBaixaPedido.colPVInfo.Count = 0 Then Exit Sub

    'Passa a Ordenaçao escolhida para o Obj
    gobjBaixaPedido.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    
    'Recarega o Grid
    lErro = ReprocessaERecarrega
    If lErro <> SUCESSO Then gError 210721
    
    Exit Sub

Erro_Ordenados_Click:

    Select Case gErr

        Case 210721

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210722)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Private Sub PedidoInicial_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoInicial_GotFocus()
        
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
    iTabPrincipalAlterado = iTabAux
    
End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 210723
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoFinal.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 210724
        End If
            
        objPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoServico_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 188828 Then gError 210725
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 210726
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:

    Cancel = True

    Select Case gErr
    
        Case 210723, 210725
        
        Case 210724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)

        Case 210726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210727)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_Change()

    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    If Len(Trim(PedidoFinal.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 210728
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoInicial.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 210729
        End If
            
        objPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoServico_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 188828 Then gError 210730
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 210731
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 210728, 210730
            
        Case 210729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)
            
        Case 210731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210732)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click
    
    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame de Pedido visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

        'Se Frame selecionado foi o de Pedido
        If TabStrip1.SelectedItem.Index = TAB_Pedidos Then
            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                Call Grid_Limpa(objGrid)
                lErro = Trata_TabPedidos()
                If lErro <> SUCESSO Then gError 210733
            End If
        End If
   
   
        Select Case iFrameAtual
        
            Case TAB_Selecao
                Parent.HelpContextID = IDH_BAIXA_PEDIDO_SELECAO
                
            Case TAB_Pedidos
                Parent.HelpContextID = IDH_BAIXA_PEDIDO_PEDIDOS
                        
        End Select
   
   End If
    
   Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 210733

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210734)

    End Select

    Exit Sub

End Sub

Function ReprocessaERecarrega() As Long

Dim lErro As Long

On Error GoTo Erro_ReprocessaERecarrega

    'Limpa a coleção de NFiscais
    Set gobjBaixaPedido = New ClassBaixaPedidos
    
    Call Move_TabSelecao_Memoria

    'Preenche a Coleção de Pedidos de Servico
    lErro = CF("BaixaPedidosSRV_ObterPedidos", gobjBaixaPedido)
    If lErro <> SUCESSO Then gError 210735

    'Limpa o GridPedido
    Call Grid_Limpa(objGrid)
        
    'Preenche o GridPedido
    Call Grid_Pedido_Preenche(gobjBaixaPedido.colPVInfo)

    ReprocessaERecarrega = SUCESSO
     
    Exit Function
    
Erro_ReprocessaERecarrega:

    ReprocessaERecarrega = gErr
     
    Select Case gErr
          
        Case 210735
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210736)
     
    End Select
     
    Exit Function

End Function

Private Function Trata_TabPedidos() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_TabPedidos

    If Ordenados.ListIndex = -1 Then
        Ordenados.ListIndex = 0
    Else
        lErro = ReprocessaERecarrega
        If lErro <> SUCESSO Then gError 210737
    End If
    
    iTabPrincipalAlterado = 0

    Exit Function

Erro_Trata_TabPedidos:

    Trata_TabPedidos = gErr

    Select Case gErr

        Case 210737

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210738)

    End Select

    Exit Function

End Function

Private Function Grid_Pedido_Preenche(colPVInfo As Collection) As Long
'Preenche o Grid Pedido com os dados de colPVInfo

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objPVInfo As ClassPVInfo
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Grid_Pedido_Preenche
    
    'Se o número de Pedidos for maior que o número de linhas do Grid
    If colPVInfo.Count + 1 > GridPedido.Rows Then

        'Altera o número de linhas do Grid de acordo com o número de Pedidos
        GridPedido.Rows = colPVInfo.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGrid)
        
    End If

    iLinha = 0

    'Percorre todas as NFiscais da Coleção
    For Each objPVInfo In colPVInfo

        iLinha = iLinha + 1

        'Passa para a tela os dados da NFiscal em questão
        GridPedido.TextMatrix(iLinha, iGrid_Baixa_Col) = objPVInfo.iMarcada
        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objPVInfo.lCodPedido
        GridPedido.TextMatrix(iLinha, iGrid_Cliente_Col) = objPVInfo.lCliente
        GridPedido.TextMatrix(iLinha, iGrid_NomeRed_Col) = objPVInfo.sClienteNomeReduzido
        If objPVInfo.dtEmissao <> DATA_NULA And objPVInfo.dtEmissao <> 0 Then GridPedido.TextMatrix(iLinha, iGrid_Emissao_Col) = Format(objPVInfo.dtEmissao, "dd/mm/yyyy")
        If objPVInfo.dtEntrega <> DATA_NULA Then GridPedido.TextMatrix(iLinha, iGrid_Entrega_Col) = Format(objPVInfo.dtEntrega, "dd/mm/yyyy")
        'Lê a Filial do Cliente para preencher Código + Nome
        objFilialCliente.lCodCliente = objPVInfo.lCliente
        objFilialCliente.iCodFilial = objPVInfo.iFilialCliente
        
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12565 Then gError 210739
                
        If lErro = 12565 Then gError 210740
        GridPedido.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    Next

    Call Grid_Refresh_Checkbox(objGrid)

    'Passa para o Obj o número de NFiscais passados pela Coleção
    objGrid.iLinhasExistentes = colPVInfo.Count

    Grid_Pedido_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Pedido_Preenche:
    
    Grid_Pedido_Preenche = gErr
    
    Select Case gErr
        
        Case 210739
        
        Case 210740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210741)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Move_TabSelecao_Memoria()
    
    'Se a DataEmissaoDe está preenchida
    If Len(Trim(DataEmissaoDe.ClipText)) > 0 Then
        gobjBaixaPedido.dtEmissaoDe = CDate(DataEmissaoDe.Text)
    'Se a DataEmissaoDe não está preenchida
    Else
        gobjBaixaPedido.dtEmissaoDe = DATA_NULA
    End If

    'Se a DataEmissaoAté está preenchida
    If Len(Trim(DataEmissaoAte.ClipText)) > 0 Then
        gobjBaixaPedido.dtEmissaoAte = CDate(DataEmissaoAte.Text)
    'Se a DataEmissaoAté não está preenchida
    Else
        gobjBaixaPedido.dtEmissaoAte = DATA_NULA
    End If

    'Se a DataEntregaDe está preenchida
    If Len(Trim(DataEntregaDe.ClipText)) > 0 Then
        gobjBaixaPedido.dtEntregaDe = CDate(DataEntregaDe.Text)
    'Se a DataEntregaDe não está preenchida
    Else
        gobjBaixaPedido.dtEntregaDe = DATA_NULA
    End If

    'Se a DataEntregaAté está preenchida
    If Len(Trim(DataEntregaAte.ClipText)) > 0 Then
        gobjBaixaPedido.dtEntregaAte = CDate(DataEntregaAte.Text)
    'Se a DataEntregaAté não está preenchida
    Else
        gobjBaixaPedido.dtEntregaAte = DATA_NULA
    End If

    'Se PedidoFinal e PedidoInicial estão preenchidos
    If Len(Trim(PedidoInicial.Text)) > 0 Then
        gobjBaixaPedido.lPedidosDe = CLng(PedidoInicial.Text)
    Else
        gobjBaixaPedido.lPedidosDe = 0
    End If
    
    If Len(Trim(PedidoFinal.Text)) > 0 Then
        gobjBaixaPedido.lPedidosAte = CLng(PedidoFinal.Text)
    Else
        gobjBaixaPedido.lPedidosAte = 0
    End If

    'Se ClienteAté e ClienteDe estão preenchidos
    If Len(Trim(ClienteDe.Text)) > 0 Then
        gobjBaixaPedido.lClientesDe = CLng(ClienteDe.Text)
    Else
        gobjBaixaPedido.lClientesDe = 0
    End If
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        gobjBaixaPedido.lClientesAte = CLng(ClienteAte.Text)
    Else
        gobjBaixaPedido.lClientesAte = 0
    End If
            
    gobjBaixaPedido.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 210742

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 210742

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210743)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 210744

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 210744

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210745)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 210746

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 33422

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210747)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 210748

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 210748

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210749)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_DownClick

    'Diminui a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 210750

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case gErr

        Case 210750

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210751)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 210752

    Exit Sub

Erro_UpDownEntregaAte_UpClick:

    Select Case gErr

        Case 210752

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210753)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_DownClick

    'Diminui a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 210754

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case gErr

        Case 210754

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210755)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 210756

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case gErr

        Case 210756

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210757)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Function Descarrega_Grid() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objPVInfo As ClassPVInfo

On Error GoTo Erro_Descarrega_Grid
    
    'Verifica se Linha se está Marcado
    For iIndice = 1 To objGrid.iLinhasExistentes
        If GridPedido.TextMatrix(iIndice, iGrid_Baixa_Col) = S_MARCADO Then
            'Se estiver
            iIndice2 = 0
            
            'Procura na coleção para excluir
            For Each objPVInfo In gobjBaixaPedido.colPVInfo
                
                'Indice para a exclusão
                iIndice2 = iIndice2 + 1
                
                If objPVInfo.lCodPedido = CLng(GridPedido.TextMatrix(iIndice, iGrid_Pedido_Col)) Then
                    
                    'Exclui da coleção global
                    gobjBaixaPedido.colPVInfo.Remove (iIndice2)
                
                End If
            Next
        End If
    Next
    
    Call Grid_Limpa(objGrid)
    
    'Preenche o GridPedido
    Call Grid_Pedido_Preenche(gobjBaixaPedido.colPVInfo)

    Descarrega_Grid = SUCESSO
     
    Exit Function
    
Erro_Descarrega_Grid:

    Descarrega_Grid = gErr
     
    Select Case gErr
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210758)
     
    End Select
     
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BAIXA_PEDIDO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Pedidos Serviço"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaPedidoSRV"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is PedidoInicial Then
            Call LabelPedidoDe_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedidoAte_Click
        ElseIf Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
End Sub

Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
End Sub

Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub


