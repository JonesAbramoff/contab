VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LiberaBloqueioGenOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   465
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CommandButton BotaoEdita 
         Caption         =   "Editar"
         Height          =   960
         Left            =   4995
         Picture         =   "LiberaBloqueioGenOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4425
         Width           =   1590
      End
      Begin VB.TextBox FilialEmpresa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4890
         TabIndex        =   34
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
         TabIndex        =   33
         Text            =   "Observacao"
         Top             =   1845
         Width           =   4245
      End
      Begin VB.CheckBox Libera 
         Height          =   210
         Left            =   195
         TabIndex        =   9
         Top             =   990
         Width           =   840
      End
      Begin VB.TextBox Codigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Text            =   "Pedido"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TipoBloqueio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   "TipoBloqueio"
         Top             =   930
         Width           =   2415
      End
      Begin VB.TextBox Usuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5595
         TabIndex        =   13
         Text            =   "Usuario"
         Top             =   960
         Width           =   1350
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera os Bloqueios Assinalados"
         Height          =   960
         Left            =   7470
         Picture         =   "LiberaBloqueioGenOcx.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4425
         Width           =   1590
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   960
         Left            =   2505
         Picture         =   "LiberaBloqueioGenOcx.ctx":0884
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4425
         Width           =   1590
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   960
         Left            =   0
         Picture         =   "LiberaBloqueioGenOcx.ctx":1A66
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4425
         Width           =   1590
      End
      Begin VB.TextBox Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   870
         TabIndex        =   14
         Text            =   "Cliente"
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox Valor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Text            =   "ValorPedido"
         Top             =   1365
         Width           =   1095
      End
      Begin MSMask.MaskEdBox DataBloqueio 
         Height          =   285
         Left            =   4140
         TabIndex        =   12
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
      Begin MSMask.MaskEdBox Data 
         Height          =   285
         Left            =   2295
         TabIndex        =   15
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
         Height          =   4110
         Left            =   15
         TabIndex        =   17
         Top             =   285
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   7250
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   555
      Width           =   9150
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Bloqueios"
         Height          =   4395
         Left            =   750
         TabIndex        =   23
         Top             =   285
         Width           =   7605
         Begin VB.Frame Frame6 
            Caption         =   "Código"
            Height          =   800
            Left            =   960
            TabIndex        =   30
            Top             =   2420
            Width           =   5520
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   300
               Left            =   780
               TabIndex        =   4
               Top             =   330
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   5
               Top             =   330
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoAte 
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
               TabIndex        =   32
               Top             =   390
               Width           =   360
            End
            Begin VB.Label LabelCodigoDe 
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
               TabIndex        =   31
               Top             =   390
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data em que foram feitos os Bloqueios"
            Height          =   800
            Left            =   960
            TabIndex        =   25
            Top             =   3300
            Width           =   5505
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   26
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
               TabIndex        =   27
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
               Index           =   0
               Left            =   330
               TabIndex        =   29
               Top             =   420
               Width           =   315
            End
            Begin VB.Label Label1 
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
               Index           =   1
               Left            =   2985
               TabIndex        =   28
               Top             =   420
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tipos"
            Height          =   2025
            Left            =   960
            TabIndex        =   24
            Top             =   315
            Width           =   5520
            Begin VB.ListBox ListaTipos 
               Columns         =   2
               Height          =   1635
               ItemData        =   "LiberaBloqueioGenOcx.ctx":2A80
               Left            =   150
               List            =   "LiberaBloqueioGenOcx.ctx":2A87
               Style           =   1  'Checkbox
               TabIndex        =   1
               Top             =   270
               Width           =   3555
            End
            Begin VB.CommandButton BotaoDesmarcarTodosTipos 
               Caption         =   "Desmarcar Todos"
               Height          =   570
               Left            =   3900
               Picture         =   "LiberaBloqueioGenOcx.ctx":2A97
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   1185
               Width           =   1425
            End
            Begin VB.CommandButton BotaoMarcarTodosTipos 
               Caption         =   "Marcar Todos"
               Height          =   570
               Left            =   3900
               Picture         =   "LiberaBloqueioGenOcx.ctx":3C79
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
      Left            =   8205
      Picture         =   "LiberaBloqueioGenOcx.ctx":4C93
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Fechar"
      Top             =   60
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   5790
      Left            =   150
      TabIndex        =   22
      Top             =   135
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   10213
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
Attribute VB_Name = "LiberaBloqueioGenOcx"
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
Dim gobjLibBloqGen As New ClassLibBloqGen
Dim gobjMapBloqGen As ClassMapeamentoBloqGen

'Grid Bloqueio:
Dim objGridBloqueio As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Codigo_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Data_Col  As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_Usuario_Col As Integer
Dim iGrid_DataBloq_Col As Integer
Dim iGrid_Libera_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Eventos de Browse
Private WithEvents objEventoCodigoDe As AdmEvento
Attribute objEventoCodigoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoAte As AdmEvento
Attribute objEventoCodigoAte.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_TELA_SELECAO = 1
Const TAB_TELA_BLOQUEIOS = 2

Private Function Traz_Bloqueios_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colTiposBloqueio As New Collection

On Error GoTo Erro_Traz_Bloqueios_Tela

    'Limpa a coleção de bloqueios
    If Not (gobjLibBloqGen.colBloqueioLiberacaoInfo Is Nothing) Then
        Do While gobjLibBloqGen.colBloqueioLiberacaoInfo.Count <> 0
            gobjLibBloqGen.colBloqueioLiberacaoInfo.Remove (1)
        Loop
    End If
      
    'Limpa o GridBloqueio
    Call Grid_Limpa(objGridBloqueio)
    
    'Verifica os Bloqueios que foram marcados
    For iIndice = 0 To ListaTipos.ListCount - 1
        If ListaTipos.Selected(iIndice) = True Then colTiposBloqueio.Add ListaTipos.ItemData(iIndice)
    Next
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("LiberacaoDeBloqueiosGen_ObterBloqueios", gobjMapBloqGen, gobjLibBloqGen, colTiposBloqueio)
    If lErro <> SUCESSO And lErro <> 29191 Then gError 198367
    
    'Preenche o GridBloqueio
    lErro = Grid_Bloqueio_Preenche(gobjLibBloqGen.colBloqueioLiberacaoInfo)
    If lErro <> SUCESSO Then gError 198368
    
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridBloqueio)
            
    Traz_Bloqueios_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = Err
    
    Select Case gErr

        Case 198367, 198368
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198369)

    End Select

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqLibInfoGen As ClassBloqLibInfoGen

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Desmarca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_INATIVO
        
        'Passa a linha do Grid para o Obj
        Set objBloqLibInfoGen = gobjLibBloqGen.colBloqueioLiberacaoInfo.Item(iLinha)
        
        'Desmarca no Obj o bloqueio em questão
        objBloqLibInfoGen.iMarcado = DESMARCADO
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqLibInfoGen As ClassBloqLibInfoGen

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Marca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO
        
        'Passa a linha do Grid para o Obj
        Set objBloqLibInfoGen = gobjLibBloqGen.colBloqueioLiberacaoInfo.Item(iLinha)
        
        'Marca no Obj o bloqueio em questão
        objBloqLibInfoGen.iMarcado = MARCADO
        
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoEditar_Click()
    
Dim lErro As Long
Dim objDocBloq As Object
Dim objBloqLibInfoGen As New ClassBloqLibInfoGen

On Error GoTo Erro_BotaoEditar_Click
    
    Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)
    
    'Verifica se alguma linha do Grid está selecionada
    If GridBloqueio.Row = 0 Then gError 198370
    
    'Passa a linha do Grid para o Obj
    Set objBloqLibInfoGen = gobjLibBloqGen.colBloqueioLiberacaoInfo.Item(GridBloqueio.Row)
    
    'Passa os dados do Bloqueio para o Obj
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objDocBloq.iFilialEmpresa = objBloqLibInfoGen.iFilialEmpresa
    End If
    
    Call CallByName(objDocBloq, gobjMapBloqGen.sNomeBrowseChave, VbLet, objBloqLibInfoGen.lCodigo)
    
    'Chama a tela de Pedidos de Venda
    Call Chama_Tela(gobjMapBloqGen.sNomeTelaEditaDocBloq, objDocBloq)
    
    Exit Sub

Erro_BotaoEditar_Click:

    Select Case gErr

        Case 198370
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198371)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridBloqueio = Nothing
    Set gobjLibBloqGen = Nothing
    Set gobjMapBloqGen = Nothing
    
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
        If lErro <> SUCESSO Then gError 198372

    End If
    
    Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198372 'Tratado na rotina chamada
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198373)

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
        If lErro <> SUCESSO Then gError 198374

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198374
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198375)

    End Select

    Exit Sub

End Sub

Private Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If
    
    colcolColecoes.Add gobjLibBloqGen.colBloqueioLiberacaoInfo
    
    Call Ordenacao_ClickGrid(objGridBloqueio, , colcolColecoes)

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

Private Sub LabelCodigoAte_Click()

Dim colSelecao As Collection
Dim objDocBloq As Object

    Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)
    
    'Passa os dados do Bloqueio para o Obj
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objDocBloq.iFilialEmpresa = giFilialEmpresa
    End If
    
    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(CodigoAte.Text)) > 0 Then
        Call CallByName(objDocBloq, gobjMapBloqGen.sNomeBrowseChave, VbLet, StrParaLong(CodigoAte.Text))
    End If
    
    Call Chama_Tela(gobjMapBloqGen.sNomeBrowseChave, colSelecao, objDocBloq, objEventoCodigoAte)

End Sub

Private Sub LabelCodigoDe_Click()

Dim colSelecao As Collection
Dim objDocBloq As Object

    Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)
    
    'Passa os dados do Bloqueio para o Obj
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objDocBloq.iFilialEmpresa = giFilialEmpresa
    End If
    
    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(CodigoDe.Text)) > 0 Then
        Call CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbLet, StrParaLong(CodigoDe.Text))
    End If
    
    Call Chama_Tela(gobjMapBloqGen.sNomeBrowseChave, colSelecao, objDocBloq, objEventoCodigoDe)

End Sub

Private Sub objEventoCodigoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDocBloq As Object
Dim bCancel As Boolean
Dim lCodigo As Long

On Error GoTo Erro_objEventoCodigoAte_evSelecao

    Set objDocBloq = obj1
    
    lCodigo = CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbGet)
    
    CodigoAte.PromptInclude = False
    CodigoAte.Text = CStr(lCodigo)
    CodigoAte.PromptInclude = True

    'Chama o Validate de CodigoAte
    Call CodigoAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoCodigoAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198376)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDocBloq As Object
Dim bCancel As Boolean
Dim lCodigo As Long

On Error GoTo Erro_objEventoCodigoDe_evSelecao

    Set objDocBloq = obj1
    
    lCodigo = CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbGet)
    
    CodigoDe.PromptInclude = False
    CodigoDe.Text = CStr(lCodigo)
    CodigoDe.PromptInclude = True
    
    'Chama o validate do CodigoDe
    Call CodigoDe_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigoDe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198377)

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

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridBloqueio = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoCodigoDe = New AdmEvento
    Set objEventoCodigoAte = New AdmEvento
    
'    'Executa a Inicialização do grid Bloqueio
'    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
'    If lErro <> SUCESSO Then gError 198378
'
'    'Limpa a Listbox ListaTipos
'    ListaTipos.Clear
'
'    'Carrega list de Bloqueios
'    lErro = TiposDeBloqueios_Carrega(ListaTipos)
'    If lErro <> SUCESSO Then gError 198379
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Case 198378, 198379

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198380)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function TiposDeBloqueios_Carrega(objListBox As ListBox) As Long
'Carrega a lista de tipos de bloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDeBloqueio As New Collection
Dim objTipoDeBloqueio As ClassTiposDeBloqueioGen

On Error GoTo Erro_TiposDeBloqueios_Carrega

    'Le todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", gobjMapBloqGen.iTipoTelaBloqueio, colTipoDeBloqueio)
    If lErro <> SUCESSO Then gError 198381

    'Preenche ListaTipos
    For Each objTipoDeBloqueio In colTipoDeBloqueio
        If objTipoDeBloqueio.iNaoApareceTelaLib = DESMARCADO Then
            objListBox.AddItem objTipoDeBloqueio.sNomeReduzido
            objListBox.ItemData(objListBox.NewIndex) = objTipoDeBloqueio.iCodigo
        End If
    Next

    TiposDeBloqueios_Carrega = SUCESSO

    Exit Function

Erro_TiposDeBloqueios_Carrega:

    TiposDeBloqueios_Carrega = gErr

    Select Case gErr

        Case 198381

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198382)

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

    If Not (gobjLibBloqGen.colCodBloqueios Is Nothing) Then
        Do While gobjLibBloqGen.colCodBloqueios.Count <> 0
            gobjLibBloqGen.colCodBloqueios.Remove (1)
        Loop
    End If
    
End Sub

Private Sub Preenche_BloqueiosCol()
'Preenche a colecao de Bloqueios

Dim iIndice As Integer

    For iIndice = 0 To ListaTipos.ListCount - 1
        If ListaTipos.Selected(iIndice) = True Then
            gobjLibBloqGen.colCodBloqueios.Add ListaTipos.ItemData(iIndice)
        End If
    Next

End Sub

Private Function Move_TabSelecao_Memoria() As Long

Dim lErro As Long
Dim lCodigoDe As Long
Dim lCodigoAte As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim iTiposSelecionados As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    dtDataDe = StrParaDate(DataDe.Text)
    dtDataAte = StrParaDate(DataAte.Text)

    'Se DataDe e DataAté estão preenchidas
    If dtDataDe <> DATA_NULA And dtDataAte <> DATA_NULA Then
        'Verifica se DataAté é maior ou igual a DataDe
        If dtDataAte < dtDataDe Then gError 198383
    End If

    'Lê CodigoDe e CodigoAte que estão na tela
    lCodigoDe = StrParaLong(CodigoDe.Text)
    lCodigoAte = StrParaLong(CodigoAte.Text)

    'Se CodigoAte e CodigoDe estão preenchidos
    If lCodigoDe <> 0 And lCodigoAte <> 0 Then
        'Verifica se CodigoAte é maior ou igual que CodigoDe
        If lCodigoAte < lCodigoDe Then gError 198384
    End If
    
    Call ContaQuant_BloqueiosSel(iTiposSelecionados)
    
    'Verifica se existe Tipo de Bloqueio selecionado
    If iTiposSelecionados < 1 Then gError 198385
    
    'Passa os dados da tela para o Obj
    gobjLibBloqGen.lCodigoAte = lCodigoAte
    gobjLibBloqGen.lCodigoDe = lCodigoDe
    gobjLibBloqGen.dtBloqueioAte = dtDataAte
    gobjLibBloqGen.dtBloqueioDe = dtDataDe
    
    'Limpa a coleção de tipos de Bloqueio selecionados
    Call Limpa_BloqueiosCol
        
    'Preenche a coleção de seleção com os tipos de bloqueio selecionados
    Call Preenche_BloqueiosCol
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function
    
Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr
    
    Select Case gErr
        
        Case 198383
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr)

        Case 198384
            Call Rotina_Erro(vbOKOnly, "ERRO_CodigoDe_MAIOR_CodigoAte", gErr)
        
        Case 198385
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_MARCADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198386)
    
    End Select
    
    Exit Function
    
End Function

Private Sub CodigoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)
End Sub

Private Sub CodigoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)
End Sub

Private Sub CodigoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_CodigoDe_Validate

    If Len(Trim(CodigoDe.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(CodigoDe.Text)
        If lErro <> SUCESSO Then gError 198387
                   
    End If
       
    Exit Sub

Erro_CodigoDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 198387
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198388)

    End Select

    Exit Sub

End Sub

Private Sub CodigoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_CodigoAte_Validate

    If Len(Trim(CodigoAte.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(CodigoAte.Text)
        If lErro <> SUCESSO Then gError 198389

    End If
       
    Exit Sub

Erro_CodigoAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 198389

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198390)

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
        If TabStripOpcao.SelectedItem.Index = TAB_TELA_BLOQUEIOS Then
            
            lErro = Move_TabSelecao_Memoria()
            If lErro <> SUCESSO Then gError 198391
            
            lErro = Traz_Bloqueios_Tela
            If lErro <> SUCESSO And lErro <> 29160 Then gError 198392
            If lErro = 29160 Then gError 198393
            
        End If
       
        Select Case iFrameAtual
        
            Case TAB_TELA_SELECAO
                Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
                
            Case TAB_TELA_BLOQUEIOS
                Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_BLOQUEIOS
                        
        End Select
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
    
        Case 198391, 198392
        
        Case 198393
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PV_SEL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198394)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198377

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 198377

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198378)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198380

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 198380

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198379)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198381

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 198381

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198382)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198383

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 198383

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198384)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(ByVal iTipoTela As Integer, Optional objDocBloq As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim objMapBloqGen As New ClassMapeamentoBloqGen
Dim sProjeto As String
Dim sClasse As String
Dim sTela As String

On Error GoTo Erro_Trata_Parametros

    objMapBloqGen.iTipoTelaBloqueio = iTipoTela
    
    lErro = CF("MapeamentoBloqGen_Le", objMapBloqGen)
    If lErro <> SUCESSO Then gError 198385
    
    Set gobjMapBloqGen = objMapBloqGen
   
    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)
    sTela = gobjMapBloqGen.sNomeTelaTestaPermissao

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then gError 198386

    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 198378
    
    'Limpa a Listbox ListaTipos
    ListaTipos.Clear
    
    'Carrega list de Bloqueios
    lErro = TiposDeBloqueios_Carrega(ListaTipos)
    If lErro <> SUCESSO Then gError 198379
    
    If Not (objDocBloq Is Nothing) Then
    
        lCodigo = CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbGet)
        
        If lCodigo > 0 Then
            CodigoDe.PromptInclude = False
            CodigoDe.Text = CStr(lCodigo)
            CodigoDe.PromptInclude = True

            CodigoAte.PromptInclude = False
            CodigoAte.Text = CStr(lCodigo)
            CodigoAte.PromptInclude = True
        End If
        
        'Marca todas as checkbox
        For iIndice = 0 To ListaTipos.ListCount - 1
            ListaTipos.Selected(iIndice) = True
        Next
        
        TabStripOpcao.Tabs.Item(TAB_TELA_BLOQUEIOS).Selected = True
        
        Call TabStripOpcao_Click
        
    End If
   
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 198378, 198379, 198385
    
        Case 198386
            Call Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198387)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Grid_Bloqueio_Preenche(colBloqLibInfoGen As Collection) As Long
'Preenche o Grid Bloqueio com os dados de colBloqLibInfoGen

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objBloqLibInfoGen As ClassBloqLibInfoGen

On Error GoTo Erro_Grid_Bloqueio_Preenche
    
    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If colBloqLibInfoGen.Count >= objGridBloqueio.objGrid.Rows Then
        Call Refaz_Grid(objGridBloqueio, colBloqLibInfoGen.Count)
    End If

    iLinha = 0

    'Percorre todos os Bloqueios da Coleção
    For Each objBloqLibInfoGen In colBloqLibInfoGen

        iLinha = iLinha + 1

        'Passa para a tela os dados do Bloqueio em questão
        If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
            GridBloqueio.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = CStr(objBloqLibInfoGen.iFilialEmpresa)
        End If
        
        GridBloqueio.TextMatrix(iLinha, iGrid_Codigo_Col) = CStr(objBloqLibInfoGen.lCodigo)
        GridBloqueio.TextMatrix(iLinha, iGrid_Cliente_Col) = objBloqLibInfoGen.sNomeRedCliForn
        GridBloqueio.TextMatrix(iLinha, iGrid_Data_Col) = Format(objBloqLibInfoGen.dtData, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_Valor_Col) = Formata_Estoque(objBloqLibInfoGen.dValor)
        GridBloqueio.TextMatrix(iLinha, iGrid_TipoBloqueio_Col) = objBloqLibInfoGen.sNomeRedTipoBloq
        GridBloqueio.TextMatrix(iLinha, iGrid_Usuario_Col) = objBloqLibInfoGen.sUsuario
        GridBloqueio.TextMatrix(iLinha, iGrid_DataBloq_Col) = Format(objBloqLibInfoGen.dtDataBloqueio, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_Observacao_Col) = objBloqLibInfoGen.sObservacao
        
    Next

    'Passa para o Obj o número de Bloqueios passados pela Coleção
    objGridBloqueio.iLinhasExistentes = colBloqLibInfoGen.Count

    Grid_Bloqueio_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Bloqueio_Preenche:
    
    Grid_Bloqueio_Preenche = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198388)
    
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
        If lErro <> SUCESSO Then gError 198389

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 198389
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198390)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio

Dim iAjuste As Integer
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Libera")
    
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objGridInt.colColuna.Add ("Filial")
    End If
    
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Tipo de Bloqueio")
    objGridInt.colColuna.Add ("Data Bloqueio")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Observação")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objGridInt.colCampo.Add (FilialEmpresa.Name)
    End If
    
    objGridInt.colCampo.Add (Codigo.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (Usuario.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    iGrid_Libera_Col = 1
    
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        iGrid_FilialEmpresa_Col = 2
        iAjuste = 1
    End If
    
    iGrid_Codigo_Col = 2 + iAjuste
    iGrid_Cliente_Col = 3 + iAjuste
    iGrid_Data_Col = 4 + iAjuste
    iGrid_Valor_Col = 5 + iAjuste
    iGrid_TipoBloqueio_Col = 6 + iAjuste
    iGrid_DataBloq_Col = 7 + iAjuste
    iGrid_Usuario_Col = 8 + iAjuste
    iGrid_Observacao_Col = 9 + iAjuste
    
    objGridInt.objGrid = GridBloqueio

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 11

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
    If lErro <> SUCESSO Then gError 198391
    
    'Descarrega o Grid de Bloqueios
    lErro = Traz_Bloqueios_Tela()
    If lErro <> SUCESSO Then gError 198392
    
    Exit Sub

Erro_BotaoLibera_Click:

    Select Case gErr

        Case 198391, 198392

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 198393)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colBloqueioGen As New Collection
Dim objBloqueioGen As New ClassBloqueioGen

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(colBloqueioGen)
    If lErro <> SUCESSO Then gError 198394
    
    If colBloqueioGen.Count = 0 Then gError 198537
    
    'Libera os Bloqueios selecionados
    lErro = CF("BloqueioGen_Libera", colBloqueioGen, gobjMapBloqGen)
    If lErro <> SUCESSO Then gError 198395
 
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 198394, 198395
        
        Case 198537
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_BLOQ_MARCADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198396)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(colBloqueioGen As Collection) As Long
'move para colBloqueioGen os bloqueios marcados para liberação  (Só move o pedido e o tipo de bloqueio pois é o suficiente)

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioGen As ClassBloqueioGen
Dim sTipoDeBloqueio As String
Dim objTipoDeBloqueio As ClassTiposDeBloqueioGen
Dim colTipoDeBloqueio As New Collection
Dim lCliente As Long
Dim dValor As Double
Dim objDocBloq As Object

On Error GoTo Erro_Move_Tela_Memoria

    'Lê todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", gobjMapBloqGen.iTipoTelaBloqueio, colTipoDeBloqueio)
    If lErro <> SUCESSO Then gError 198397

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes
        
        'se o elemento está marcado para ser liberado
        If GridBloqueio.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objBloqueioGen = New ClassBloqueioGen
            
            sTipoDeBloqueio = GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col)
            
            'pega o codigo do tipo de bloqueio
            For Each objTipoDeBloqueio In colTipoDeBloqueio
            
                If objTipoDeBloqueio.sNomeReduzido = sTipoDeBloqueio Then
                    objBloqueioGen.iTipoDeBloqueio = objTipoDeBloqueio.iCodigo
                    Exit For
                End If
            
            Next
            
            If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
                objBloqueioGen.iFilialEmpresa = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_FilialEmpresa_Col))
            End If
            objBloqueioGen.lCodigo = StrParaLong(GridBloqueio.TextMatrix(iIndice, iGrid_Codigo_Col))
            
            Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)
                
            'Passa os dados do Bloqueio para o Obj
            If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
                objDocBloq.iFilialEmpresa = objBloqueioGen.iFilialEmpresa
            End If
                
            Call CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbLet, objBloqueioGen.lCodigo)
            
            lErro = CF(gobjMapBloqGen.sNomeFuncLeDoc, objDocBloq)
            If lErro <> SUCESSO Then gError 198538
            
            lCliente = objDocBloq.lCliente
            dValor = CallByName(objDocBloq, gobjMapBloqGen.sClasseDocNomeValor, VbGet)
            
            objBloqueioGen.sCodUsuarioLib = gsUsuario
            objBloqueioGen.sResponsavelLib = gsUsuario
            objBloqueioGen.dtDataLib = gdtDataAtual
            objBloqueioGen.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
            objBloqueioGen.lClienteDoc = lCliente
            objBloqueioGen.dValorDoc = dValor
            
            colBloqueioGen.Add objBloqueioGen
            
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 198397, 198538
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198398)
            
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
    Name = "LiberaBloqueioGen"
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
Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub TabStripOpcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripOpcao)
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridBloqueio)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Sub BotaoEdita_Click()

Dim lErro As Long
Dim objDocBloq As Object
Dim objBloqLibInfoGen As ClassBloqLibInfoGen

On Error GoTo Erro_BotaoEdita_Click

    If GridBloqueio.Row = 0 Then gError 198535

    Set objBloqLibInfoGen = gobjLibBloqGen.colBloqueioLiberacaoInfo.Item(GridBloqueio.Row)
    
    Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)
        
    'Passa os dados do Bloqueio para o Obj
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objDocBloq.iFilialEmpresa = giFilialEmpresa
    End If
        
    Call CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbLet, objBloqLibInfoGen.lCodigo)

    'Lê o Documento
    lErro = Chama_Tela(gobjMapBloqGen.sNomeTelaEditaDocBloq, objDocBloq)
    If lErro <> SUCESSO Then gError 198536
    
    Exit Sub

Erro_BotaoEdita_Click:

    Select Case gErr
    
        Case 198535
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 198536

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 198393)

    End Select

    Exit Sub
    
End Sub
