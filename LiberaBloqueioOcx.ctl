VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LiberaBloqueioOcx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8310
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   795
      Visible         =   0   'False
      Width           =   7755
      Begin VB.TextBox FilialEmpresa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4890
         TabIndex        =   37
         Text            =   "FilialEmpresa"
         Top             =   1755
         Width           =   615
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   615
         MaxLength       =   250
         TabIndex        =   36
         Text            =   "Observacao"
         Top             =   1845
         Width           =   4245
      End
      Begin VB.CheckBox Libera 
         Height          =   210
         Left            =   195
         TabIndex        =   10
         Top             =   990
         Width           =   840
      End
      Begin VB.TextBox Pedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1065
         TabIndex        =   11
         Text            =   "Pedido"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TipoBloqueio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "TipoBloqueio"
         Top             =   930
         Width           =   2415
      End
      Begin VB.TextBox Usuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5595
         TabIndex        =   14
         Text            =   "Usuario"
         Top             =   960
         Width           =   1350
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   156
         Width           =   4725
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera os Bloqueios Assinalados"
         Height          =   960
         Left            =   165
         Picture         =   "LiberaBloqueioOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3420
         Width           =   1590
      End
      Begin VB.CommandButton BotaoPedido 
         Caption         =   "Pedido de Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   22
         Top             =   4020
         Width           =   1725
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   3945
         Picture         =   "LiberaBloqueioOcx.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3615
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   2325
         Picture         =   "LiberaBloqueioOcx.ctx":1624
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3615
         Width           =   1425
      End
      Begin VB.TextBox Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   870
         TabIndex        =   15
         Text            =   "Cliente"
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox ValorPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Text            =   "ValorPedido"
         Top             =   1365
         Width           =   1095
      End
      Begin MSMask.MaskEdBox DataBloqueio 
         Height          =   285
         Left            =   4140
         TabIndex        =   13
         Top             =   960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DataPedido 
         Height          =   285
         Left            =   2295
         TabIndex        =   16
         Top             =   1350
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
         Height          =   2640
         Left            =   75
         TabIndex        =   18
         Top             =   615
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   4657
         _Version        =   393216
         Rows            =   11
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
         Left            =   135
         TabIndex        =   35
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   825
      Width           =   7770
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Bloqueios"
         Height          =   4305
         Left            =   750
         TabIndex        =   25
         Top             =   75
         Width           =   6270
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   800
            Left            =   435
            TabIndex        =   32
            Top             =   2420
            Width           =   5520
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   780
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
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   3450
               TabIndex        =   5
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
               Top             =   390
               Width           =   360
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
               TabIndex        =   33
               Top             =   390
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data em que foram feitos os Bloqueios"
            Height          =   800
            Left            =   435
            TabIndex        =   27
            Top             =   3300
            Width           =   5505
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   367
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   795
               TabIndex        =   6
               Top             =   367
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
               Left            =   3390
               TabIndex        =   7
               Top             =   367
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
               Left            =   4560
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   360
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               TabIndex        =   31
               Top             =   420
               Width           =   315
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
               TabIndex        =   30
               Top             =   420
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tipos"
            Height          =   2025
            Left            =   435
            TabIndex        =   26
            Top             =   315
            Width           =   5520
            Begin VB.ListBox ListaTipos 
               Columns         =   2
               Height          =   1635
               ItemData        =   "LiberaBloqueioOcx.ctx":263E
               Left            =   150
               List            =   "LiberaBloqueioOcx.ctx":2645
               Style           =   1  'Checkbox
               TabIndex        =   1
               Top             =   270
               Width           =   3555
            End
            Begin VB.CommandButton BotaoDesmarcarTodosTipos 
               Caption         =   "Desmarcar Todos"
               Height          =   570
               Left            =   3900
               Picture         =   "LiberaBloqueioOcx.ctx":2655
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   1185
               Width           =   1425
            End
            Begin VB.CommandButton BotaoMarcarTodosTipos 
               Caption         =   "Marcar Todos"
               Height          =   570
               Left            =   3900
               Picture         =   "LiberaBloqueioOcx.ctx":3837
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   465
               Width           =   1425
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
      Left            =   6840
      Picture         =   "LiberaBloqueioOcx.ctx":4851
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   4935
      Left            =   150
      TabIndex        =   24
      Top             =   405
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   8705
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
Attribute VB_Name = "LiberaBloqueioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer 'apenas p/uso pela interface c/grid
Dim iFrameAtual As Integer
Dim gobjLiberacaoBloqueios As New ClassLiberacaoBloqueios
Dim asOrdenacao(4) As String
Dim asOrdenacaoString(4) As String

'Grid Bloqueio:
Dim objGridBloqueio As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_DataPedido_Col  As Integer
Dim iGrid_ValorPedido_col As Integer
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_Usuario_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Libera_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Eventos de Browse
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_Selecao = 1
Const TAB_BLOQUEIOS = 2


Private Function Traz_Bloqueios_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colTiposBloqueio As New Collection

On Error GoTo Erro_Traz_Bloqueios_Tela

    'Limpa a coleção de bloqueios
    If Not (gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo Is Nothing) Then
        
        Do While gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Count <> 0
        
            gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Remove (1)
        
        Loop
        
    End If
      
    'Limpa o GridBloqueio
    Call Grid_Limpa(objGridBloqueio)
    
    'Verifica os Bloqueios que foram marcados
    For iIndice = 0 To ListaTipos.ListCount - 1
                    
        If ListaTipos.Selected(iIndice) = True Then colTiposBloqueio.Add ListaTipos.ItemData(iIndice)
        
    Next
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("LiberacaoDeBloqueios_ObterBloqueios", gobjLiberacaoBloqueios, colTiposBloqueio)
    If lErro <> SUCESSO And lErro <> 29191 Then Error 29189
    If lErro = 29191 Then Error 29160
    
    'Preenche o GridBloqueio
    lErro = Grid_Bloqueio_Preenche(gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo)
    If lErro <> SUCESSO Then Error 58160
    
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridBloqueio)
            
    Traz_Bloqueios_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = Err
    
    Select Case Err

        Case 29189, 29160
            
        Case 58160
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162320)

    End Select

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Desmarca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_INATIVO
        
        'Passa a linha do Grid para o Obj
        Set objBloqueioLiberacaoInfo = gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Item(iLinha)
        
        'Desmarca no Obj o bloqueio em questão
        objBloqueioLiberacaoInfo.iMarcado = DESMARCADO
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Marca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO
        
        'Passa a linha do Grid para o Obj
        Set objBloqueioLiberacaoInfo = gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Item(iLinha)
        
        'Marca no Obj o bloqueio em questão
        objBloqueioLiberacaoInfo.iMarcado = MARCADO
        
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoPedido_Click()
    
Dim objBloqueioLiberacaoInfo As New ClassBloqueioLiberacaoInfo
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_BotaoPedido_Click
    
    'Verifica se alguma linha do Grid está selecionada
    If GridBloqueio.Row = 0 Then Error 41604
    
    'Passa a linha do Grid para o Obj
    Set objBloqueioLiberacaoInfo = gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Item(GridBloqueio.Row)
    
    'Passa os dados do Bloqueio para o Obj
    objPedidoDeVenda.iFilialEmpresa = objBloqueioLiberacaoInfo.iFilialEmpresa
    objPedidoDeVenda.lCodigo = objBloqueioLiberacaoInfo.lCodPedido
    
    'Chama a tela de Pedidos de Venda
    Call Chama_Tela("PedidoVenda", objPedidoDeVenda)
    
    Exit Sub

Erro_BotaoPedido_Click:

    Select Case Err

        Case 41604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162321)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridBloqueio = Nothing
    Set gobjLiberacaoBloqueios = Nothing
    
End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Se a DataAte está preenchida
    If Len(DataAte.ClipText) > 0 Then

        'Verifica se a DataAte é válida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then Error 41606

        If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub
        
        'Verifica se a Datade é Menor que a DataAte
        If CDate(DataDe.Text) > CDate(DataAte.Text) Then Error 58299

    End If
    
    Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 41606 'Tratado na rotina chamada
        
        Case 58299
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162322)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe está preenchida
    If Len(DataDe.ClipText) > 0 Then

        'Verifica se a DataDe é válida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then Error 29162

        If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

        'Verifica se a Datade é Menor que a DataAte
        If CDate(DataDe.Text) > CDate(DataAte.Text) Then Error 58300

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 29162
        
        Case 58300
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162323)

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

Private Sub LabelPedidoAte_Click()

Dim colSelecao As Collection
Dim objPedidoDeVenda As New ClassPedidoDeVenda

    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(PedidoFinal.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoFinal.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoAte)

End Sub

Private Sub LabelPedidoDe_Click()

Dim colSelecao As Collection
Dim objPedidoDeVenda As New ClassPedidoDeVenda

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(PedidoInicial.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoInicial.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoDe)

End Sub


Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim bCancel As Boolean

On Error GoTo Erro_objEventoPedidoAte_evSelecao

    Set objPedidoDeVenda = obj1
    
    PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)

    'Chama o Validate de PedidoFinal
    Call PedidoFinal_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoPedidoAte_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162324)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim bCancel As Boolean

On Error GoTo Erro_objEventoPedidoDe_evSelecao

    Set objPedidoDeVenda = obj1
    
    PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)
    
    'Chama o validate do PedidoInicial
    Call PedidoInicial_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoPedidoDe_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162325)

    End Select

    Exit Sub

End Sub

Private Sub Ordenados_Click()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Ordenados_Click
    
    'Verifica se a coleção de Bloqueios está vazia
    If gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Count <> 0 Then
    
        'Passa a Ordenaçao escolhida para o Obj
        gobjLiberacaoBloqueios.sOrdenacao = asOrdenacao(Ordenados.ItemData(Ordenados.ListIndex))
            
        'Recarrega o Grid com a Nova ordenacao
        lErro = Traz_Bloqueios_Tela
        If lErro <> SUCESSO And lErro <> 29160 Then Error 41605
        
        If lErro = 29160 Then Error 41607
        
    End If
            
    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 41605
        
        Case 41607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PV_SEL", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162326)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoDesmarcarTodosTipos_Click()
'Desmarca todas as checkbox da ListBox Bloqueios

Dim iIndice As Integer

    'Percorre todas as checkbox da ListaTipos
    For iIndice = 0 To ListaTipos.ListCount - 1

        'Desmarca na tela o bloqueio em questão
        ListaTipos.Selected(iIndice) = False

    Next

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoMarcarTodosTipos_Click()
'Marca todas as checkbox da ListBox Bloqueios

Dim iIndice As Integer
    
    'Percorre todas as checkbox da ListaTipos
    For iIndice = 0 To ListaTipos.ListCount - 1
            
        'Marca na tela o bloqueio em questão
        ListaTipos.Selected(iIndice) = True
        
    Next
    
End Sub


Private Sub Ordenacao_Carrega()
'preenche a combo de ordenacao e inicializa variaveis globais

Dim iIndice As Integer

    'Cria vetor de Ordenacao (BD, TELA)
    asOrdenacao(0) = "BloqueiosPV.PedidoDeVenda"
    asOrdenacao(1) = "BloqueiosPV.CodUsuario, BloqueiosPV.PedidoDeVenda"
    asOrdenacao(2) = "PedidosDeVenda.DataEmissao, PedidosDeVenda.Codigo"
    asOrdenacao(3) = "BloqueiosPV.Data, BloqueiosPV.PedidoDeVenda"
    asOrdenacao(4) = "BloqueiosPV.TipoDeBloqueio, BloqueiosPV.PedidoDeVenda"
    
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

    'Seleciona a opção TipoDeBloqueio + PedidoDeVendas de seleção
    Ordenados.ListIndex = 4

End Sub

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridBloqueio = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 29163
    
    'Limpa a Listbox ListaTipos
    ListaTipos.Clear
    
    'Carrega list de Bloqueios
    lErro = TiposDeBloqueios_Carrega(ListaTipos)
    If lErro <> SUCESSO Then gError 29184
    
    lErro = CF("LiberacaoBloqueio_FilialEmpresa", gobjLiberacaoBloqueios)
    If lErro <> SUCESSO Then gError 126952
    
    'preenche a combo de ordenacao
    Call Ordenacao_Carrega
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 29163, 29184, 126952

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162327)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function TiposDeBloqueios_Carrega(objListBox As ListBox) As Long
'Carrega a lista de tipos de bloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDeBloqueio As New Collection
Dim objTipoDeBloqueio As ClassTipoDeBloqueio

On Error GoTo Erro_TiposDeBloqueios_Carrega

    'Le todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueio_Le_Todos", colTipoDeBloqueio)
    If lErro <> SUCESSO And lErro <> 29168 Then Error 29164

    'Preenche ListaTipos
    For Each objTipoDeBloqueio In colTipoDeBloqueio
    
        If objTipoDeBloqueio.sNomeReduzido <> "Bloqueio Parcial" And objTipoDeBloqueio.sNomeReduzido <> "Não Reserva" Then
            objListBox.AddItem objTipoDeBloqueio.sNomeReduzido
            objListBox.ItemData(objListBox.NewIndex) = objTipoDeBloqueio.iCodigo
        End If
        
    Next

    TiposDeBloqueios_Carrega = SUCESSO

    Exit Function

Erro_TiposDeBloqueios_Carrega:

    TiposDeBloqueios_Carrega = Err

    Select Case Err

        Case 29164

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162328)

    End Select

    Exit Function

End Function

Private Sub ContaQuant_BloqueiosSel(iTiposSelecionados As Integer)

Dim iIndice As Integer

    iTiposSelecionados = 0
    
    'Conta a quantidade de Tipos de Bloqueio selecionados
    For iIndice = 0 To ListaTipos.ListCount - 1
                    
        If ListaTipos.Selected(iIndice) = True Then iTiposSelecionados = iTiposSelecionados + 1
        
    Next

End Sub

Private Sub Limpa_BloqueiosCol()

    If Not (gobjLiberacaoBloqueios.colCodBloqueios Is Nothing) Then
        
        Do While gobjLiberacaoBloqueios.colCodBloqueios.Count <> 0
        
            gobjLiberacaoBloqueios.colCodBloqueios.Remove (1)
        
        Loop
        
    End If
    
End Sub

Private Sub Preenche_BloqueiosCol()
'Preenche a colecao de Bloqueios

Dim iIndice As Integer

    For iIndice = 0 To ListaTipos.ListCount - 1
            
        If ListaTipos.Selected(iIndice) = True Then
        
            gobjLiberacaoBloqueios.colCodBloqueios.Add ListaTipos.ItemData(iIndice)
        
        End If
        
    Next

End Sub

Private Function Move_TabSelecao_Memoria() As Long

Dim lErro As Long
Dim lPedidoInicial As Long
Dim lPedidoFinal As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim iTiposSelecionados As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    'Se a DataDe está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then
        dtDataDe = CDate(DataDe.Text)
    'Se a DataDe não está preenchida
    Else
        dtDataDe = DATA_NULA
    End If

    'Se a DataAté está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then
        dtDataAte = CDate(DataAte.Text)
    'Se a DataAté não está preenchida
    Else
        dtDataAte = DATA_NULA
    End If

    'Se DataDe e DataAté estão preenchidas
    If dtDataDe <> DATA_NULA And dtDataAte <> DATA_NULA Then

        'Verifica se DataAté é maior ou igual a DataDe
        If dtDataAte < dtDataDe Then Error 29180

    End If

    'Se PedidoFinal e PedidoInicial estão preenchidos
    If Len(Trim(PedidoFinal.Text)) > 0 And Len(Trim(PedidoInicial.Text)) > 0 Then

        'Lê PedidoInicial e PedidoFinal que estão na tela
        lPedidoInicial = CLng(Trim(PedidoInicial.Text))
        lPedidoFinal = CLng(Trim(PedidoFinal.Text))

        'Verifica se PedidoFinal é maior ou igual que PedidoInicial
        If lPedidoFinal < lPedidoInicial Then Error 29181

    End If
    
    Call ContaQuant_BloqueiosSel(iTiposSelecionados)
    
    'Verifica se existe Tipo de Bloqueio selecionado
    If iTiposSelecionados < 1 Then Error 29190
    
    'Passa os dados da tela para o Obj
    gobjLiberacaoBloqueios.lPedVendasAte = lPedidoFinal
    gobjLiberacaoBloqueios.lPedVendasDe = lPedidoInicial
    gobjLiberacaoBloqueios.dtBloqueioAte = dtDataAte
    gobjLiberacaoBloqueios.dtBloqueioDe = dtDataDe
    gobjLiberacaoBloqueios.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    
    'Limpa a coleção de tipos de Bloqueio selecionados
    Call Limpa_BloqueiosCol
        
    'Preenche a coleção de seleção com os tipos de bloqueio selecionados
    Call Preenche_BloqueiosCol
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function
    
Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err
    
    Select Case Err
        
        Case 29180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", Err)

        Case 29181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
        
        Case 29190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_MARCADO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162329)
    
    End Select
    
    Exit Function
    
End Function

Private Sub PedidoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)

End Sub

Private Sub PedidoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)

End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 58295
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoFinal.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 58296
        End If
        
        objPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 64394
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 64395
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:

    Cancel = True

    Select Case gErr
    
        Case 58295, 64394
        
        Case 58296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)
        
        Case 64395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162330)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    If Len(Trim(PedidoFinal.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 58297
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoInicial.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 58298
        End If
        
        objPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 64391
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 64392
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 58297, 64391
            
        Case 58298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)
            
        Case 64392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162331)

    End Select

    Exit Sub

End Sub

Private Sub TabStripOpcao_Click()

Dim lErro As Long

On Error GoTo Erro_TabStripOpcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStripOpcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStripOpcao, Me) <> SUCESSO Then Exit Sub
       
        'Torna Frame de Bloqueios visível
        Frame1(TabStripOpcao.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStripOpcao.SelectedItem.Index
       
        'Se Frame selecionado foi o de Bloqueios
        If TabStripOpcao.SelectedItem.Index = TAB_BLOQUEIOS Then
            
            lErro = Move_TabSelecao_Memoria()
            If lErro <> SUCESSO Then Error 22978
            
            lErro = Traz_Bloqueios_Tela
            If lErro <> SUCESSO And lErro <> 29160 Then Error 29183
            If lErro = 29160 Then Error 29171
            
        End If
       
        Select Case iFrameAtual
        
            Case TAB_Selecao
                Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
                
            Case TAB_BLOQUEIOS
                Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_BLOQUEIOS
                        
        End Select
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case Err

        Case 22978
            'Limpa o GridBloqueio
            Call Grid_Limpa(objGridBloqueio)
        
        Case 29171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PV_SEL", Err)
'            TabStripOpcao.Tabs(iFrameAtual).Selected = True
        
        Case 29183
'            Set TabStripOpcao.SelectedItem = TabStripOpcao.Tabs(iFrameAtual)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162332)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 41608

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case Err

        Case 41608

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162333)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29172

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case Err

        Case 29172

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162334)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29173

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case Err

        Case 29173

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162335)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29174

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case Err

        Case 29174

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162336)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPedidoDeVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoDeVenda Is Nothing) Then
        
        If objPedidoDeVenda.lCodigo > 0 Then PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)
        If objPedidoDeVenda.lCodigo > 0 Then PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)
        
        'Marca todas as checkbox
        For iIndice = 0 To ListaTipos.ListCount - 1
            
            ListaTipos.Selected(iIndice) = True
        
        Next
                
        'selecionar o tab p/liberacao dos bloqueios do pedido
'        If iFrameAtual <> 2 Then
'            ''''Set TabStripOpcao.SelectedItem = TabStripOpcao.Tabs(2)
'            TabStripOpcao.Tabs.Item(2).Selected = True
'        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162337)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Grid_Bloqueio_Preenche(colBloqueioLiberacaoInfo As Collection) As Long
'Preenche o Grid Bloqueio com os dados de colBloqueioLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

On Error GoTo Erro_Grid_Bloqueio_Preenche

    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If colBloqueioLiberacaoInfo.Count + 1 > GridBloqueio.Rows Then
    
        If colBloqueioLiberacaoInfo.Count > NUM_MAX_BLOQUEIOS_LIBERACAO Then Error 19166

        'Altera o número de linhas do Grid de acordo com o número de Bloqueios
        GridBloqueio.Rows = colBloqueioLiberacaoInfo.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridBloqueio)

    End If

    iLinha = 0

    'Percorre todos os Bloqueios da Coleção
    For Each objBloqueioLiberacaoInfo In colBloqueioLiberacaoInfo

        iLinha = iLinha + 1

        'Passa para a tela os dados do Bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = CStr(objBloqueioLiberacaoInfo.iFilialEmpresa)
        GridBloqueio.TextMatrix(iLinha, iGrid_Pedido_Col) = CStr(objBloqueioLiberacaoInfo.lCodPedido)
        GridBloqueio.TextMatrix(iLinha, iGrid_Cliente_Col) = objBloqueioLiberacaoInfo.sNomeReduzidoCliente
        GridBloqueio.TextMatrix(iLinha, iGrid_DataPedido_Col) = Format(objBloqueioLiberacaoInfo.dtDataEmissao, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_ValorPedido_col) = Formata_Estoque(objBloqueioLiberacaoInfo.dValorPedido)
        GridBloqueio.TextMatrix(iLinha, iGrid_TipoBloqueio_Col) = objBloqueioLiberacaoInfo.sNomeReduzidoTipoBloqueio
        GridBloqueio.TextMatrix(iLinha, iGrid_Usuario_Col) = objBloqueioLiberacaoInfo.sUsuario
        GridBloqueio.TextMatrix(iLinha, iGrid_Data_Col) = Format(objBloqueioLiberacaoInfo.dtDataBloqueio, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_Observacao_Col) = objBloqueioLiberacaoInfo.sObservacao
        
    Next

    'Passa para o Obj o número de Bloqueios passados pela Coleção
    objGridBloqueio.iLinhasExistentes = colBloqueioLiberacaoInfo.Count

    Grid_Bloqueio_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Bloqueio_Preenche:
    
    Grid_Bloqueio_Preenche = Err
    
    Select Case Err
    
        Case 19166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_BLOQUEIO_MAIOR_LIMITE", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162338)
    
    End Select
    
    Exit Function
    
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 29182

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 29182
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162339)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Libera")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Data Pedido")
    objGridInt.colColuna.Add ("Valor Pedido")
    objGridInt.colColuna.Add ("Tipo de Bloqueio")
    objGridInt.colColuna.Add ("Data Bloqueio")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Observacao")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (DataPedido.Name)
    objGridInt.colCampo.Add (ValorPedido.Name)
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (Usuario.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    iGrid_Libera_Col = 1
    iGrid_FilialEmpresa_Col = 2
    iGrid_Pedido_Col = 3
    iGrid_Cliente_Col = 4
    iGrid_DataPedido_Col = 5
    iGrid_ValorPedido_col = 6
    iGrid_TipoBloqueio_Col = 7
    iGrid_Data_Col = 8
    iGrid_Usuario_Col = 9
    iGrid_Observacao_Col = 10
    
    objGridInt.objGrid = GridBloqueio

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'todas as linhas do grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'largura da primeira coluna
    GridBloqueio.ColWidth(0) = 500

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Private Sub BotaoLibera_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLibera_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 36641
    
    'Descarrega o Grid de Bloqueios
    lErro = Descarrega_Grid()
    If lErro <> SUCESSO Then Error 36709
    
    Exit Sub

Erro_BotaoLibera_Click:

    Select Case Err

        Case 36641, 36709

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162340)

    End Select

    Exit Sub

End Sub

Function Descarrega_Grid() As Long
'Descarrega o Grid com os Bloqueios selecionados

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo
Dim objTipoBloqueio As New ClassTipoDeBloqueio

On Error GoTo Erro_Descarrega_Grid
    
    Call Grid_Limpa(objGridBloqueio)
    
    'Preenche o Grid de Bloqueios
    Call Grid_Bloqueio_Preenche(gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo)

    Descarrega_Grid = SUCESSO
     
    Exit Function
    
Erro_Descarrega_Grid:

    Descarrega_Grid = Err
     
    Select Case Err
          
        Case 58309 'Tratado na Rotina chamada
        
        Case 58310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", Err, objTipoBloqueio.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162341)
     
    End Select
     
    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colBloqueioPV As New Collection
Dim objBloqueioPV As New ClassBloqueioPV

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(colBloqueioPV)
    If lErro <> SUCESSO Then Error 36644
    
    'Libera os Bloqueios selecionados
    lErro = CF("Bloqueio_Libera", colBloqueioPV)
    If lErro <> SUCESSO Then Error 36668

    lErro = Atualiza_gobjLiberacaoBloqueios(colBloqueioPV)
    If lErro <> SUCESSO Then Error 60776
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 36644, 36668, 60776

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162342)

    End Select

    Exit Function

End Function

Private Function Atualiza_gobjLiberacaoBloqueios(colBloqueioPV As Collection) As Long

Dim objBloqueioPV As ClassBloqueioPV
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_gobjLiberacaoBloqueios

    For Each objBloqueioPV In colBloqueioPV
    
        If objBloqueioPV.dtDataLib <> DATA_NULA Then
    
            For iIndice = 1 To gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Count
    
                Set objBloqueioLiberacaoInfo = gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Item(iIndice)
    
                If objBloqueioPV.lPedidoDeVendas = objBloqueioLiberacaoInfo.lCodPedido And _
                   objBloqueioPV.iSequencial = objBloqueioLiberacaoInfo.iSeqBloqueio Then
                   
                    gobjLiberacaoBloqueios.colBloqueioLiberacaoInfo.Remove (iIndice)
                    
                    Exit For
                   
                End If
    
            Next

        End If

    Next

    Atualiza_gobjLiberacaoBloqueios = SUCESSO
    
    Exit Function
    
Erro_Atualiza_gobjLiberacaoBloqueios:

    Atualiza_gobjLiberacaoBloqueios = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162343)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria(colBloqueioPV As Collection) As Long
'move para colBloqueioPV os bloqueios marcados para liberação  (Só move o pedido e o tipo de bloqueio pois é o suficiente)

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPV As ClassBloqueioPV
Dim sTipoDeBloqueio As String
Dim objTipoDeBloqueio As ClassTipoDeBloqueio
Dim colTipoDeBloqueio As New Collection

On Error GoTo Erro_Move_Tela_Memoria

    'Lê todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueio_Le_Todos", colTipoDeBloqueio)
    If lErro <> SUCESSO And lErro <> 29168 Then Error 36645

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes
        
        'se o elemento está marcado para ser liberado
        If GridBloqueio.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objBloqueioPV = New ClassBloqueioPV
            
            sTipoDeBloqueio = GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col)
            
            'pega o codigo do tipo de bloqueio
            For Each objTipoDeBloqueio In colTipoDeBloqueio
            
                If objTipoDeBloqueio.sNomeReduzido = sTipoDeBloqueio Then
                    objBloqueioPV.iTipoDeBloqueio = objTipoDeBloqueio.iCodigo
                    Exit For
                End If
            
            Next
            
            objBloqueioPV.iFilialEmpresa = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_FilialEmpresa_Col))
            objBloqueioPV.lPedidoDeVendas = StrParaLong(GridBloqueio.TextMatrix(iIndice, iGrid_Pedido_Col))
            objBloqueioPV.sResponsavel = gsUsuario
            objBloqueioPV.sCodUsuario = gsUsuario
            objBloqueioPV.sCodUsuarioLib = gsUsuario
            objBloqueioPV.sResponsavelLib = gsUsuario
            objBloqueioPV.dtDataLib = gdtDataAtual
            objBloqueioPV.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
            
            colBloqueioPV.Add objBloqueioPV
            
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
    
        Case 36645
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162344)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Liberação de Bloqueio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LiberaBloqueio"
    
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


Private Sub TabStripOpcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripOpcao)
End Sub

