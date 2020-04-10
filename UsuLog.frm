VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form UsuLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuários Logados"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7785
      ScaleHeight     =   480
      ScaleWidth      =   1050
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   1110
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "UsuLog.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Height          =   360
         Left            =   75
         Picture         =   "UsuLog.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Atualizar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox Login 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   6615
      MaxLength       =   64
      TabIndex        =   4
      Top             =   2520
      Width           =   1965
   End
   Begin VB.TextBox Computador 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   4800
      MaxLength       =   64
      TabIndex        =   3
      Top             =   2565
      Width           =   1965
   End
   Begin VB.TextBox CodGrupo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3030
      MaxLength       =   64
      TabIndex        =   2
      Top             =   2595
      Width           =   1965
   End
   Begin VB.TextBox CodUsuario 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1245
      MaxLength       =   64
      TabIndex        =   1
      Top             =   2640
      Width           =   1965
   End
   Begin MSFlexGridLib.MSFlexGrid GridUsu 
      Height          =   4410
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   7779
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Usuários logados no sistema:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   345
      Width           =   5700
   End
End
Attribute VB_Name = "UsuLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAlterado As Integer
Dim objGridUsu As AdmGrid

Dim iGrid_CodUsuario_Col As Integer
Dim iGrid_CodGrupo_Col As Integer
Dim iGrid_Computador_Col As Integer
Dim iGrid_Login_Col As Integer

Private Sub BotaoAtualizar_Click()
    Call Traz_Usuarios_Logados
End Sub

Private Sub GridUsu_Click()
Dim iExecutaEntradaCelula As Integer
    Call Grid_Click(objGridUsu, iExecutaEntradaCelula)
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridUsu, iAlterado)
    End If
End Sub

Private Sub GridUsu_GotFocus()
    Call Grid_Recebe_Foco(objGridUsu)
End Sub

Private Sub GridUsu_EnterCell()
    Call Grid_Entrada_Celula(objGridUsu, iAlterado)
End Sub

Private Sub GridUsu_LeaveCell()
    If objGridUsu.iSaidaCelula = 1 Then Call Saida_Celula(objGridUsu)
End Sub

Private Sub GridUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridUsu)
End Sub

Private Sub GridUsu_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer
    Call Grid_Trata_Tecla(KeyAscii, objGridUsu, iExecutaEntradaCelula)
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridUsu, iAlterado)
    End If
End Sub

Private Sub GridUsu_LostFocus()
    Call Grid_Libera_Foco(objGridUsu)
End Sub

Private Sub GridUsu_RowColChange()
    Call Grid_RowColChange(objGridUsu)
End Sub

Private Sub GridUsu_Scroll()
    Call Grid_Scroll(objGridUsu)
End Sub

Private Function Grid_Inicia(ByVal objGrid1 As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Grid_Inicia

    'tela em questão
    Set objGrid1.objForm = Me

    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Cód. Usuário")
    objGrid1.colColuna.Add ("Cód. Grupo")
    objGrid1.colColuna.Add ("Computador")
    objGrid1.colColuna.Add ("Login PC")

   'campos de edição do grid
    objGrid1.colCampo.Add (CodUsuario.Name)
    objGrid1.colCampo.Add (CodGrupo.Name)
    objGrid1.colCampo.Add (Computador.Name)
    objGrid1.colCampo.Add (Login.Name)
    
    iGrid_CodUsuario_Col = 1
    iGrid_CodGrupo_Col = 2
    iGrid_Computador_Col = 3
    iGrid_Login_Col = 4

    'Grid
    objGrid1.objGrid = GridUsu

    objGrid1.iProibidoExcluir = 1
    objGrid1.iProibidoIncluir = 1

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 16

    'todas as linhas do grid
    objGrid1.objGrid.Rows = 501

    objGrid1.objGrid.ColWidth(0) = 400

    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid1)

    Grid_Inicia = SUCESSO

    Exit Function

Erro_Grid_Inicia:

    Grid_Inicia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161724)

    End Select

    Exit Function

End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGridUsu = New AdmGrid
    
    'Inicializa Grid
    lErro = Grid_Inicia(objGridUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Traz_Usuarios_Logados

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161726)

    End Select

    Exit Sub

End Sub

Private Sub Traz_Usuarios_Logados()

Dim lErro As Long, iIndice As Integer
Dim colUsu As New Collection
Dim objUsu As ClassDicUsuario

On Error GoTo Erro_Traz_Usuarios_Logados

    Call Grid_Limpa(objGridUsu)
    
    lErro = Usuarios_Le_Todos1(colUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iIndice = 0
    For Each objUsu In colUsu
        If objUsu.iLogado = MARCADO Then
            iIndice = iIndice + 1
            GridUsu.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objUsu.sCodUsuario
            GridUsu.TextMatrix(iIndice, iGrid_Computador_Col) = objUsu.sComputador
            GridUsu.TextMatrix(iIndice, iGrid_Login_Col) = objUsu.sNomeLogin
            GridUsu.TextMatrix(iIndice, iGrid_CodGrupo_Col) = objUsu.sCodGrupo
        End If
    Next
    
    objGridUsu.iLinhasExistentes = iIndice
    
    Exit Sub

Erro_Traz_Usuarios_Logados:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161726)

    End Select

    Exit Sub

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError ERRO_SEM_MENSAGEM

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161729)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub
