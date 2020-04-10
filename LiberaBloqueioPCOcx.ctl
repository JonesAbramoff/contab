VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LiberaBloqueioPCOcx 
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
      Height          =   8100
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   765
      Visible         =   0   'False
      Width           =   16455
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera os Bloqueios Assinalados"
         Height          =   960
         Left            =   360
         Picture         =   "LiberaBloqueioPCOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7035
         Width           =   1590
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "LiberaBloqueioPCOcx.ctx":0442
         Left            =   1680
         List            =   "LiberaBloqueioPCOcx.ctx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   150
         Width           =   3330
      End
      Begin VB.TextBox Responsavel 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6450
         TabIndex        =   29
         Text            =   "Responsável"
         Top             =   2460
         Width           =   3000
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Pedido de Compra ..."
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
         Left            =   6645
         TabIndex        =   28
         Top             =   7545
         Width           =   2010
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   600
         Left            =   2925
         Picture         =   "LiberaBloqueioPCOcx.ctx":0446
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7215
         Width           =   1440
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   600
         Left            =   4740
         Picture         =   "LiberaBloqueioPCOcx.ctx":1460
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7215
         Width           =   1440
      End
      Begin VB.TextBox ValorPedido 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5085
         TabIndex        =   25
         Text            =   "ValorPedido"
         Top             =   2925
         Width           =   1350
      End
      Begin VB.TextBox Usuario 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Text            =   "Usuario"
         Top             =   2445
         Width           =   2000
      End
      Begin VB.TextBox TipoBloqueio 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "Tipo de Bloqueio"
         Top             =   2400
         Width           =   2000
      End
      Begin VB.TextBox Pedido 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Text            =   "Pedido"
         Top             =   2430
         Width           =   1095
      End
      Begin VB.CheckBox Libera 
         Height          =   210
         Left            =   360
         TabIndex        =   21
         Top             =   2130
         Width           =   840
      End
      Begin VB.TextBox Fornecedor 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   390
         TabIndex        =   18
         Text            =   "Fornecedor"
         Top             =   2880
         Width           =   3000
      End
      Begin VB.TextBox Filial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2070
         TabIndex        =   17
         Text            =   "Filial"
         Top             =   2880
         Width           =   1065
      End
      Begin MSMask.MaskEdBox DataPedido 
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   2910
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DataBloqueio 
         Height          =   285
         Left            =   3675
         TabIndex        =   20
         Top             =   2445
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
         Height          =   6225
         Left            =   300
         TabIndex        =   31
         Top             =   765
         Width           =   15900
         _ExtentX        =   28046
         _ExtentY        =   10980
         _Version        =   393216
         Rows            =   11
         Cols            =   7
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
         Left            =   300
         TabIndex        =   32
         Top             =   210
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   765
      Width           =   16440
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Bloqueios"
         Height          =   4740
         Left            =   510
         TabIndex        =   1
         Top             =   315
         Width           =   6270
         Begin VB.Frame Frame4 
            Caption         =   "Tipos"
            Height          =   2000
            Left            =   435
            TabIndex        =   14
            Top             =   315
            Width           =   5520
            Begin VB.CommandButton BotaoDesselecTodos 
               Caption         =   "Desmarcar Todos"
               Height          =   600
               Left            =   2925
               Picture         =   "LiberaBloqueioPCOcx.ctx":2642
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1170
               Width           =   1440
            End
            Begin VB.CommandButton BotaoSelecionaTodos 
               Caption         =   "Marcar Todos"
               Height          =   600
               Left            =   1110
               Picture         =   "LiberaBloqueioPCOcx.ctx":3824
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   1170
               Width           =   1440
            End
            Begin VB.ListBox ListaTipos 
               Columns         =   2
               Height          =   735
               ItemData        =   "LiberaBloqueioPCOcx.ctx":483E
               Left            =   885
               List            =   "LiberaBloqueioPCOcx.ctx":4845
               Style           =   1  'Checkbox
               TabIndex        =   15
               Top             =   285
               Width           =   3720
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data em que foram feitos os Bloqueios"
            Height          =   800
            Left            =   435
            TabIndex        =   7
            Top             =   3300
            Width           =   5505
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   810
               TabIndex        =   8
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
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1965
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   375
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   10
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
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   4590
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   375
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
               TabIndex        =   13
               Top             =   428
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
               TabIndex        =   12
               Top             =   428
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   800
            Left            =   435
            TabIndex        =   2
            Top             =   2420
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   990
               TabIndex        =   3
               Top             =   330
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
               Left            =   3690
               TabIndex        =   4
               Top             =   330
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
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
               Left            =   615
               TabIndex        =   6
               Top             =   390
               Width           =   315
            End
            Begin VB.Label Label6 
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
               Left            =   3255
               TabIndex        =   5
               Top             =   390
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
      Left            =   15495
      Picture         =   "LiberaBloqueioPCOcx.ctx":4855
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Fechar"
      Top             =   90
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8595
      Left            =   150
      TabIndex        =   34
      Top             =   405
      Width           =   16620
      _ExtentX        =   29316
      _ExtentY        =   15161
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueios"
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
Attribute VB_Name = "LiberaBloqueioPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjLiberacaoBloqueiosPC As New ClassLiberacaoBloqueiosPC
Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iTabSelecaoAlterado As Integer
Dim asOrdenacao(4) As String
Dim asOrdenacaoString(4) As String

'grid Bloqueio:
Dim objGridBloqueio As AdmGrid
Dim iGrid_Libera_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_DataPedido_Col  As Integer
Dim iGrid_ValorPedido_col As Integer
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_Usuario_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_Responsavel_Col As Integer

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 49329

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 49329
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162345)

    End Select

    Exit Function

End Function

Private Sub DataAte_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    iTabSelecaoAlterado = iTabSelecao

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
' Critica a data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Se a DataAte está preenchida
    If Len(DataAte.ClipText) > 0 Then

        'Verifica se a DataAte é valida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then Error 49218

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True
    
    Select Case Err

        Case 49218

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162346)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    iTabSelecaoAlterado = iTabSelecao

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Critica a data

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe está preenchida
    If Len(DataDe.ClipText) > 0 Then

        'Verifica se a DataDe é valida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then Error 49217

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True
    
    Select Case Err

        Case 49217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162347)

    End Select

    Exit Sub

End Sub

Private Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Private Sub GridBloqueio_GotFocus()
    Call Grid_Recebe_Foco(objGridBloqueio)
End Sub

Private Sub GridBloqueio_EnterCell()
    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
End Sub

Private Sub GridBloqueio_LeaveCell()
    Call Saida_Celula(objGridBloqueio)
End Sub

Private Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)
End Sub

Private Sub GridBloqueio_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Private Sub GridBloqueio_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridBloqueio)
End Sub

Private Sub GridBloqueio_RowColChange()
    Call Grid_RowColChange(objGridBloqueio)
End Sub

Private Sub GridBloqueio_Scroll()
    Call Grid_Scroll(objGridBloqueio)
End Sub

Private Function Move_TabSelecao_Memoria() As Long

Dim lErro As Long
Dim lPedidoInicial As Long '??? Por que criou essas vars? Poderia atribuir direto no gobj.
Dim lPedidoFinal As Long '??? Por que criou essas vars?
Dim dtDataDe As Date '??? Por que criou essas vars?
Dim dtDataAte As Date '??? Por que criou essas vars?
Dim iTiposSelecionados As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria


    dtDataDe = StrParaDate(DataDe.Text)
    dtDataAte = StrParaDate(DataAte.Text)
    
    'Se DataDe e DataAté estão preenchidas
    If dtDataDe <> DATA_NULA And dtDataAte <> DATA_NULA Then

        'Verifica se DataAté é maior ou igual a DataDe
        If dtDataAte < dtDataDe Then Error 49238

    End If

    'Le PedidoInicial e PedidoFinal que estão na tela
    lPedidoInicial = StrParaLong(PedidoInicial.Text)
    lPedidoFinal = StrParaLong(PedidoFinal.Text)

    
    'Se PedidoFinal e PedidoInicial estão preenchidos
    If Len(Trim(PedidoFinal.Text)) > 0 And Len(Trim(PedidoInicial.Text)) > 0 Then
        'Verifica se PedidoFinal é maior ou igual que PedidoInicial
        If lPedidoFinal < lPedidoInicial Then Error 49239
    End If

    iTiposSelecionados = 0

    'Conta a quantidade de Tipos de Bloqueio selecionados
    For iIndice = 0 To ListaTipos.ListCount - 1
        If ListaTipos.Selected(iIndice) = True Then iTiposSelecionados = iTiposSelecionados + 1
    Next

    'Verifica se existe Tipo de Bloqueio selecionado
    If iTiposSelecionados < 1 Then Error 49240

    'Passa os dados da tela para o Obj
    gobjLiberacaoBloqueiosPC.lPedComprasAte = StrParaLong(PedidoFinal.Text)
    gobjLiberacaoBloqueiosPC.lPedComprasDe = StrParaLong(PedidoInicial.Text)
    gobjLiberacaoBloqueiosPC.dtBloqueioAte = dtDataAte
    gobjLiberacaoBloqueiosPC.dtBloqueioDe = dtDataDe
    gobjLiberacaoBloqueiosPC.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    'Limpa a coleção de tipos de Bloqueio selecionados
    Call Limpa_BloqueiosCol

    'Preenche a coleção de seleção com os tipos de bloqueio selecionados
    Call Preenche_BloqueiosCol

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err

    Select Case Err

        Case 49238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)
            
        Case 49239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)

        Case 49240
             lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIOPC_NAO_MARCADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162348)

    End Select

    Exit Function

End Function

Private Sub Limpa_BloqueiosCol()
'Limpa a coleção de tipos de Bloqueio selecionados

    If Not (gobjLiberacaoBloqueiosPC.colTipoBloqueio Is Nothing) Then

        Do While gobjLiberacaoBloqueiosPC.colTipoBloqueio.Count <> 0

            gobjLiberacaoBloqueiosPC.colTipoBloqueio.Remove (1)

        Loop

    End If

End Sub
Private Sub Preenche_BloqueiosCol()
'Preenche a coleção de seleção com os tipos de bloqueio selecionados

Dim iIndice As Integer

    For iIndice = 0 To ListaTipos.ListCount - 1

        If ListaTipos.Selected(iIndice) = True Then

            gobjLiberacaoBloqueiosPC.colTipoBloqueio.Add ListaTipos.ItemData(iIndice)

        End If

    Next

End Sub

Private Function Traz_Bloqueios_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colTiposBloqueioPC As New Collection

On Error GoTo Erro_Traz_Bloqueios_Tela

    'Limpa a coleção de bloqueios
    If Not (gobjLiberacaoBloqueiosPC.colBloqueioPC Is Nothing) Then

        Do While gobjLiberacaoBloqueiosPC.colBloqueioPC.Count <> 0

            gobjLiberacaoBloqueiosPC.colBloqueioPC.Remove (1)

        Loop

    End If

    'Limpa o GridBloqueio
    Call Grid_Limpa(objGridBloqueio)

    'Verifica os Bloqueios que foram marcados
    For iIndice = 0 To ListaTipos.ListCount - 1
        If ListaTipos.Selected(iIndice) = True Then colTiposBloqueioPC.Add ListaTipos.ItemData(iIndice)
    Next

    'Preenche a Coleção de Bloqueios
    lErro = CF("LiberacaoDeBloqueiosPC_ObterBloqueios", gobjLiberacaoBloqueiosPC, colTiposBloqueioPC)
    If lErro <> SUCESSO And lErro <> 51180 Then Error 49235
    If lErro = 51180 Then Error 49326
    
    'Preenche o GridBloqueio
     Call Grid_Bloqueio_Preenche(gobjLiberacaoBloqueiosPC.colBloqueioPC)

    'Atualiza As checkboxes
    Call Grid_Refresh_Checkbox(objGridBloqueio)

    Traz_Bloqueios_Tela = SUCESSO

    Exit Function

Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = Err

    Select Case Err

        Case 49235, 49326

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162349)

    End Select

    Exit Function

End Function

Private Sub Grid_Bloqueio_Preenche(colBloqueioPCLiberacaoInfo As Collection)
'Preenche o Grid Bloqueio com os dados de colBloqueioPCLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objBloqueioPC  As ClassBloqueioPC
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Grid_Bloqueio_Preenche

    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If colBloqueioPCLiberacaoInfo.Count + 1 > GridBloqueio.Rows Then

        If colBloqueioPCLiberacaoInfo.Count > NUM_MAX_BLOQUEIOSPC_LIBERACAO Then Error 49255

        'Altera o número de linhas do Grid de acordo com o número de Bloqueios
        GridBloqueio.Rows = colBloqueioPCLiberacaoInfo.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Inicializa_Grid_Bloqueio(objGridBloqueio)

    End If

    iLinha = 0
    
    'Percorre todos os Bloqueios da Coleção
    For Each objBloqueioPC In colBloqueioPCLiberacaoInfo

        iLinha = iLinha + 1

        'Passa para a tela os dados do Bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Pedido_Col) = objBloqueioPC.lPedCompras
        GridBloqueio.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objBloqueioPC.sNomeReduzidoFornecedor
        objFilialEmpresa.iCodFilial = objBloqueioPC.iFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 57249
        GridBloqueio.TextMatrix(iLinha, iGrid_Filial_Col) = objBloqueioPC.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
        GridBloqueio.TextMatrix(iLinha, iGrid_DataPedido_Col) = Format(objBloqueioPC.dtData, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_ValorPedido_col) = Format(objBloqueioPC.dValorPedido, "Standard")
        GridBloqueio.TextMatrix(iLinha, iGrid_TipoBloqueio_Col) = objBloqueioPC.iTipoBloqueio & SEPARADOR & objBloqueioPC.sNomeReduzidoTipoBloqueio
        GridBloqueio.TextMatrix(iLinha, iGrid_Usuario_Col) = objBloqueioPC.sCodUsuario
        GridBloqueio.TextMatrix(iLinha, iGrid_DataBloqueio_Col) = Format(objBloqueioPC.dtData, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_Responsavel_Col) = objBloqueioPC.sResponsavel

    Next

    'Passa para o Obj o número de BloqueiosPC passados pela Coleção
    objGridBloqueio.iLinhasExistentes = colBloqueioPCLiberacaoInfo.Count

    Exit Sub

Erro_Grid_Bloqueio_Preenche:

    Select Case Err

        Case 49255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_BLOQUEIOS_SELECIONADOS_SUPERIOR_MAXIMO", Err)

        Case 57249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162350)

    End Select

    Exit Sub

End Sub


Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoCompra Is Nothing) Then

        PedidoInicial.Text = CStr(objPedidoCompra.lCodigo)
        PedidoFinal.Text = CStr(objPedidoCompra.lCodigo)

        'Marca todas as checkbox
        For iIndice = 0 To ListaTipos.ListCount - 1

            ListaTipos.Selected(iIndice) = True

        Next

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162351)

    End Select

    Exit Function

End Function

Private Function TiposDeBloqueiosPC_Carrega(objListBox As ListBox) As Long
'Carrega a listbox  de tipos de bloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDeBloqueioPC  As New Collection
Dim objTipoDeBloqueioPC  As ClassTipoBloqueioPC
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_TiposDeBloqueiosPC_Carrega

    'Preenche a listbox TiposdeBloqueio
    'Le cada codigo e Nome Reduzido da tabela TiposdeBloqueioPC
    lErro = CF("Cod_Nomes_Le", "TiposdeBloqueioPC", "Codigo", "NomeReduzido", STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 49232

    'Preenche ListaTipos
    For Each objCodigoNome In colCodigoNome

        objListBox.AddItem objCodigoNome.sNome
        objListBox.ItemData(objListBox.NewIndex) = objCodigoNome.iCodigo

    Next

    TiposDeBloqueiosPC_Carrega = SUCESSO

    Exit Function

Erro_TiposDeBloqueiosPC_Carrega:

    TiposDeBloqueiosPC_Carrega = Err

    Select Case Err

        Case 49232

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162352)

    End Select

    Exit Function

End Function

 Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Bloqueio

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Libera")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data do Pedido")
    objGridInt.colColuna.Add ("Valor do Pedido")
    objGridInt.colColuna.Add ("Tipo de Bloqueio")
    objGridInt.colColuna.Add ("Data do Bloqueio")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")

    ' campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (DataPedido.Name)
    objGridInt.colCampo.Add (ValorPedido.Name)
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (Usuario.Name)
    objGridInt.colCampo.Add (Responsavel.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Libera_Col = 1
    iGrid_Pedido_Col = 2
    iGrid_Fornecedor_Col = 3
    iGrid_Filial_Col = 4
    iGrid_DataPedido_Col = 5
    iGrid_ValorPedido_col = 6
    iGrid_TipoBloqueio_Col = 7
    iGrid_DataBloqueio_Col = 8
    iGrid_Usuario_Col = 9
    iGrid_Responsavel_Col = 10

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridBloqueio

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOSPC_LIBERACAO + 1

    'Não permite incluir novas linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 18

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Bloqueio:

    Inicializa_Grid_Bloqueio = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162353)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioPC As ClassBloqueioPC

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Desmarca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_INATIVO

        'Passa a linha do Grid para o Obj
        Set objBloqueioPC = gobjLiberacaoBloqueiosPC.colBloqueioPC.Item(iLinha)

        'Desmarca no Obj o bloqueio em questão
        objBloqueioPC.iMarcado = DESMARCADO

    Next

    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)

End Sub

Private Sub BotaoDesselecTodos_Click()
'Desmarca todas as checkbox da ListBox de Tipos de Bloqueio

Dim iIndice As Integer

    'Percorre todas as checkbox da ListaTipos
    For iIndice = 0 To ListaTipos.ListCount - 1

        'Desmarca na tela o bloqueio em questão
        ListaTipos.Selected(iIndice) = False

    Next

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLibera_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLibera_Click

    'Chama rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 49223

    lErro = Traz_Bloqueios_Tela()
    If lErro <> SUCESSO Then Error 49283

    Exit Sub

Erro_BotaoLibera_Click:

    Select Case Err

        Case 49223, 49283

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162354)

    End Select

    Exit Sub

End Sub
Private Function Move_Tela_Memoria(colBloqueiosPC As Collection) As Long
'move para colBloqueiosPC os bloqueios marcados para liberação

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPC  As ClassBloqueioPC
Dim sTipoDeBloqueioPC As String
Dim objTipoDeBloqueioPC As ClassTipoBloqueioPC
Dim colTipoDeBloqueioPC As New Collection

On Error GoTo Erro_Move_Tela_Memoria

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        'se o elemento está marcado para ser liberado
        If GridBloqueio.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then

            Set objBloqueioPC = gobjLiberacaoBloqueiosPC.colBloqueioPC.Item(iIndice)

            'Informa o usuario e a data do sistema
            objBloqueioPC.sCodUsuarioLib = gsUsuario
            objBloqueioPC.dtDataLib = gdtDataAtual

            colBloqueiosPC.Add objBloqueioPC

        End If

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162355)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colBloqueiosPC As New Collection
On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Tela_Memoria(colBloqueiosPC)
    If lErro <> SUCESSO Then Error 49227

    'Se nao ha bloqueios no grid selecionados para liberacao
    If colBloqueiosPC.Count = 0 Then Error 49228

    lErro = CF("BloqueiosPC_Libera", colBloqueiosPC)
    If lErro <> SUCESSO Then Error 49229

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 49227, 49229

        Case 49228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_BLOQUEIOS_LIBERAR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162356)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub BotaoMarcarTodos_Click()
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioPC As ClassBloqueioPC

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Marca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO

        'Passa a linha do Grid para o Obj
        Set objBloqueioPC = gobjLiberacaoBloqueiosPC.colBloqueioPC.Item(iLinha)

        'Marca no Obj o bloqueio em questão
        objBloqueioPC.iMarcado = MARCADO

    Next

    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)


End Sub

Private Sub BotaoPedido_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim lErro As Long

On Error GoTo Erro_BotaoPedido_Click

    'Verifica se alguma linha do Grid está selecionada
    If GridBloqueio.Row = 0 Then Error 49226

    'Passa os dados do Bloqueio para o Obj
    objPedidoCompra.iFilialEmpresa = gobjLiberacaoBloqueiosPC.colBloqueioPC.Item(GridBloqueio.Row).iFilialEmpresa
    objPedidoCompra.lCodigo = gobjLiberacaoBloqueiosPC.colBloqueioPC.Item(GridBloqueio.Row).lPedCompras

    'Chama a tela de Pedidos de Compra
    Call Chama_Tela("PedComprasCons", objPedidoCompra)

    Exit Sub

Erro_BotaoPedido_Click:

    Select Case Err

        Case 49226

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162357)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSelecionaTodos_Click()
'Marca todas as checkbox da ListBox Tipos de Bloqueios

Dim iIndice As Integer

    'Percorre todas as checkbox da Lista Tipos
    For iIndice = 0 To ListaTipos.ListCount - 1

        'Marca na tela o bloqueio em questão
        ListaTipos.Selected(iIndice) = True

    Next

End Sub

Private Sub DataAte_Change()

    iTabSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_Change()

    iTabSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objGridBloqueio = New AdmGrid
    Set gobjLiberacaoBloqueiosPC = New ClassLiberacaoBloqueiosPC

    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then Error 49224

    'Limpa a Listbox ListaTipos
    ListaTipos.Clear

    'Carrega list de Bloqueios
    lErro = TiposDeBloqueiosPC_Carrega(ListaTipos)
    If lErro <> SUCESSO Then Error 49225

    'preenche a combo de ordenacao
    Call Ordenacao_Carrega

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 49224, 49225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162358)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Ordenacao_Carrega()
'preenche a combo de ordenacao e inicializa variaveis globais

Dim iIndice As Integer

    'Carregar os arrays de ordenação dos Bloqueios
    asOrdenacao(0) = "BloqueiosPC.PedCompras"
    asOrdenacao(1) = "BloqueiosPC.CodUsuario, BloqueiosPC.PedCompras"
    asOrdenacao(2) = "PedidoCompra.DataEmissao, PedidoCompra.Codigo"
    asOrdenacao(3) = "BloqueiosPC.Data, BloqueiosPC.PedCompras"
    asOrdenacao(4) = "BloqueiosPC.TipoDeBloqueio,BloqueiosPC.PedCompras"

    asOrdenacaoString(0) = "Pedido"
    asOrdenacaoString(1) = "Usuário + Pedido"
    asOrdenacaoString(2) = "Data de Emissão do Pedido + Pedido"
    asOrdenacaoString(3) = "Data do Bloqueio + Pedido"
    asOrdenacaoString(4) = "Tipo de Bloqueio + Pedido"

    'Carrega a Combobox Ordenados
    For iIndice = 0 To 4

        Ordenados.AddItem asOrdenacaoString(iIndice)
        Ordenados.ItemData(Ordenados.NewIndex) = iIndice

    Next

    'Seleciona a opção TipoDeBloqueio + PedidoDeCompra de seleção
    Ordenados.ListIndex = 4

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGridBloqueio = Nothing
    Set gobjLiberacaoBloqueiosPC = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub ListaTipos_Click()

    iTabSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Ordenados_Click

    'Verifica se a coleção de Bloqueios está vazia
    If gobjLiberacaoBloqueiosPC.colBloqueioPC.Count <> 0 Then

        'Passa a Ordenaçao escolhida para o Obj
        gobjLiberacaoBloqueiosPC.sOrdenacao = asOrdenacao(Ordenados.ItemData(Ordenados.ListIndex))

        lErro = Traz_Bloqueios_Tela()
        If lErro <> SUCESSO And lErro <> 49326 Then Error 49233
        If lErro = 49326 Then Error 49328
        
    End If

    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 49233
        
        Case 49328
             lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PC_SEL", Err)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162359)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_Change()

    iTabSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoFinal_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
End Sub

Private Sub PedidoInicial_Change()

    iTabSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PedidoInicial_GotFocus()

Dim iTabSelecao As Integer

    iTabSelecao = iTabSelecaoAlterado
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
    iTabSelecaoAlterado = iTabSelecao
    
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
    
    'Se Frame selecionado foi o de Bloqueios
    If TabStrip1.SelectedItem.Index = 1 Then iTabSelecaoAlterado = 0
    
    If TabStrip1.SelectedItem.Index = 2 And iTabSelecaoAlterado = REGISTRO_ALTERADO Then

        lErro = Move_TabSelecao_Memoria()
        If lErro <> SUCESSO Then Error 49236

        lErro = Traz_Bloqueios_Tela()
        If lErro <> SUCESSO And lErro <> 49326 Then Error 49237
        If lErro = 49326 Then Error 49327
        
    End If
        
    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case 49236, 49237

        Case 49327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PC_SEL", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162360)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49219

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case Err

        Case 49219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162361)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um dia em DataAte
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49220

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case Err

        Case 49220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162362)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49221

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case Err

        Case 49221

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162363)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49222

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case Err

        Case 49222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162364)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Liberação de Bloqueios  -  Pedidos de Compra"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LiberaBloqueioPC"
    
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



Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

