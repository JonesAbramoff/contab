VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GrupoRotina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo x Rotina"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   1245
   ClientWidth     =   9480
   Icon            =   "GrupoRotina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7635
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "GrupoRotina.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "GrupoRotina.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GrupoRotina.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton SelecionaTodas 
      Caption         =   "Marcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1545
      Picture         =   "GrupoRotina.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      Width           =   1830
   End
   Begin VB.CommandButton DesselecionarTodas 
      Caption         =   "Desmarcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3750
      Picture         =   "GrupoRotina.frx":196E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4485
      Width           =   1830
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   3675
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   390
      Width           =   2670
   End
   Begin VB.TextBox Classe 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6375
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3345
      Width           =   1650
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "GrupoRotina.frx":2B50
      Left            =   3450
      List            =   "GrupoRotina.frx":2B5A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3300
      Width           =   1275
   End
   Begin VB.TextBox Projeto 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4725
      MaxLength       =   50
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3345
      Width           =   1650
   End
   Begin VB.ComboBox Grupo 
      Height          =   315
      Left            =   1095
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid GridRotinas 
      Height          =   1890
      Left            =   435
      TabIndex        =   2
      Top             =   855
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
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
      Height          =   210
      Left            =   2850
      TabIndex        =   7
      Top             =   435
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   435
      Width           =   615
   End
End
Attribute VB_Name = "GrupoRotina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAlterado As Integer
Dim objGrid1 As AdmGrid

Private Sub DesselecionarTodas_Click()

Dim iRow As Integer
    
    For iRow = 1 To objGrid1.iLinhasExistentes
        GridRotinas.TextMatrix(iRow, 1) = "Sem Acesso"
    Next

End Sub

Private Sub GridRotinas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If


End Sub

Private Sub GridRotinas_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridRotinas_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridRotinas_LeaveCell()

    If objGrid1.iSaidaCelula = 1 Then Call Saida_Celula(objGrid1)

End Sub

Private Sub GridRotinas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Private Sub GridRotinas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridRotinas_LostFocus()

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridRotinas_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridRotinas_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Private Function GrupoRotina_Exibe(ByVal sGrupo As String, ByVal sModulo As String) As Long
'Exibe na Tela os dados de GrupoRotinas correspondentes ao Grupo e ao Módulo (rotinas dentro do módulo)

Dim lErro As Long
Dim colGrupoRotina As New colGrupoRotina
Dim objGrupoRotina As ClassDicGrupoRotina
Dim iRow As Integer, iNovoNumRows As Integer

On Error GoTo Erro_GrupoRotina_Exibe

    'Lê dados de GrupoRotinas correspondentes ao Grupo e ao Módulo
    lErro = GrupoRotina_Le_GrupoModulo(sGrupo, sModulo, colGrupoRotina)
    If lErro Then Error 6440

    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoRotina.Count

    iNovoNumRows = IIf(colGrupoRotina.Count > objGrid1.iLinhasVisiveis, colGrupoRotina.Count, objGrid1.iLinhasVisiveis) + 1

    If iNovoNumRows <> GridRotinas.Rows Then
    
        GridRotinas.Rows = iNovoNumRows
        
        're-inicializa o grid
        Call Grid_Inicializa(objGrid1)
    
    End If
    
    'Coloca dados de GrupoRotinas no Grid
    iRow = GridRotinas.FixedRows - 1
    For Each objGrupoRotina In colGrupoRotina
        iRow = iRow + 1
        GridRotinas.TextMatrix(iRow, 0) = objGrupoRotina.sSiglaRotina
        GridRotinas.TextMatrix(iRow, 1) = IIf(objGrupoRotina.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridRotinas.TextMatrix(iRow, 2) = objGrupoRotina.sProjeto
        GridRotinas.TextMatrix(iRow, 3) = objGrupoRotina.sClasse
    Next

    GrupoRotina_Exibe = SUCESSO

    Exit Function

Erro_GrupoRotina_Exibe:

    GrupoRotina_Exibe = Err

    Select Case Err

        Case 6440  'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161730)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_Local()

Dim iLinha As Integer
    
    'Limpa ComboBox Grupo
    Grupo.ListIndex = -1
    
    'Limpa GridRotinas
    If objGrid1.iLinhasExistentes > 0 Then
        
        Call Grid_Limpa(objGrid1)
        
        'Limpa a primeira coluna
        For iLinha = 1 To GridRotinas.Rows - 1
            GridRotinas.TextMatrix(iLinha, 0) = ""
        Next

    End If

End Function

Private Function Grid_Inicia() As Long

Dim lErro As Long
Dim colRotina As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Grid_Inicia

    'Cria AdmGrid
    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    'titulos do grid
    objGrid1.colColuna.Add ("Rotina")
    objGrid1.colColuna.Add ("Tipo de Acesso")
    objGrid1.colColuna.Add ("Projeto Customizado")
    objGrid1.colColuna.Add ("Classe Customizada")

   'campos de edição do grid
    objGrid1.colCampo.Add (Acesso.Name)
    objGrid1.colCampo.Add (Projeto.Name)
    objGrid1.colCampo.Add (Classe.Name)

    'Grid
    objGrid1.objGrid = GridRotinas

    'Proibe a exclusão de linhas do GRID
    objGrid1.iProibidoExcluir = 1

    'Proibe a inclusão de linhas no GRID
    objGrid1.iProibidoIncluir = 1

    'Lê os Grupos existentes no BD
    lErro = Rotinas_Le_Todas(colRotina)
    If lErro Then Error 6441

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9

    'todas as linhas do grid
    GridRotinas.Rows = IIf(colRotina.Count > objGrid1.iLinhasVisiveis, colRotina.Count, objGrid1.iLinhasVisiveis) + GridRotinas.FixedRows

    GridRotinas.ColWidth(0) = 3900

    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGrid1)

    'Limpa os índices na primeira coluna
    For iIndice = 1 To GridRotinas.Rows - 1

        GridRotinas.TextMatrix(iIndice, 0) = ""

    Next

    'Alinha a primeira coluna à esquerda
    GridRotinas.ColAlignment(0) = flexAlignLeftCenter

    Grid_Inicia = SUCESSO

    Exit Function

Erro_Grid_Inicia:

    Grid_Inicia = Err

    Select Case Err

        Case 6441   'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161731)

    End Select

    Exit Function


End Function

Private Sub Acesso_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Acesso_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Acesso_LostFocus()

    Set objGrid1.objControle = Acesso
    Call Grid_Campo_Libera_Foco(objGrid1)

End Sub

Private Sub BotaoFechar_Click()

    Unload GrupoRotina

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim colGrupoRotina As New colGrupoRotina
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Rotina foram informados
    If Len(Grupo.Text) = 0 Then Error 6442

    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 6443

    'Armazena linhas do Grid em colGrupoRotina
    For iIndice = 1 To objGrid1.iLinhasExistentes
        colGrupoRotina.Add LOG_NAO, GridRotinas.TextMatrix(iIndice, 2), GridRotinas.TextMatrix(iIndice, 3), Grupo.Text, GridRotinas.TextMatrix(iIndice, 0), IIf(GridRotinas.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO)
    Next

    'Grava colGrupoRotina no banco de dados (é um update)
    lErro = GrupoRotina_Grava2(Modulo.Text, colGrupoRotina)
    If lErro Then Error 6444

    'Limpa a Tela
    Call Limpa_Tela_Local

Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 6442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_INFORMADO", Err)
            Grupo.SetFocus

        Case 6443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID_ROTINAS", Err)

        Case 6444  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161732)

     End Select

     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    'Limpa a Tela
    Call Limpa_Tela_Local

End Sub

Private Sub Classe_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Classe_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Classe_LostFocus()

    Set objGrid1.objControle = Classe
    Call Grid_Campo_Libera_Foco(objGrid1)

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colGrupo As New Collection
Dim vGrupo As Variant
Dim colModulo As New Collection
Dim vModulo As Variant

On Error GoTo Erro_GrupoRotina_Form_Load

    Me.HelpContextID = IDH_GRUPO_ROTINA
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 6445
    
    'Lê Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 6446
    If lErro Then Error 6447

    'Preenche List da ComboBox Grupo
    For Each vGrupo In colGrupo
        Grupo.AddItem (vGrupo)
    Next

    'Verifica há Grupo selecionado
    If Len(gsGrupo) > 0 Then
    
        'Seleciona gsGrupo na ComboBox Grupo
        Call ListBox_Select(gsGrupo, Grupo)
        gsGrupo = ""
        
    Else  'Não há grupo selecionado
    
        'Seleciona primeiro Grupo na ComboBox
        Grupo.ListIndex = 0
        
    End If
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6448

    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next

    'Seleciona o primeiro Módulo na ComboBox Modulo
    Modulo.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_GrupoRotina_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err

        Case 6445, 6447, 6448  'Tratado na rotina chamada
        
        Case 6446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161733)

    End Select

    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Modulo_Click
    
    If Len(Grupo.Text) > 0 Then
        
        'Limpa o Grid
        If objGrid1.iLinhasExistentes > 0 Then
            
            Call Grid_Limpa(objGrid1)
            
            'Limpa a primeira coluna
            For iLinha = 1 To GridRotinas.Rows - 1
                GridRotinas.TextMatrix(iLinha, 0) = ""
            Next

        End If

        'Preenche Grid com dados de GrupoRotinas associados ao Grupo e às rotinas do Módulo
        lErro = GrupoRotina_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then Error 6449
        
    End If

    Exit Sub

Erro_Modulo_Click:

    Select Case Err

        Case 6449  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161734)

    End Select

    Exit Sub

End Sub

Private Sub Projeto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Projeto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Projeto_LostFocus()

    Set objGrid1.objControle = Projeto
    Call Grid_Campo_Libera_Foco(objGrid1)

End Sub

Private Sub Grupo_Click()

Dim lErro As Long

On Error GoTo Erro_Grupo_Click

    If Len(Grupo.Text) > 0 And Len(Modulo.Text) > 0 Then
    
        'Preenche Grid com dados de GrupoRotinas associados ao Grupo e às rotinas do Módulo
        lErro = GrupoRotina_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then Error 6450
        
    End If
    
    Exit Sub

Erro_Grupo_Click:

    Select Case Err

        Case 6450  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161735)

    End Select

    Exit Sub

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case 1
                Set objGridInt.objControle = Acesso

                'critica da coluna 1 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6451

            Case 2
                Set objGridInt.objControle = Projeto

                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6452

            Case 3
                Set objGridInt.objControle = Classe

                'Critica da coluna 3 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6453

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 6454

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 6451, 6452, 6453, 6454
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161736)

    End Select

    Exit Function

End Function

Private Sub SelecionaTodas_Click()

Dim iRow As Integer

    For iRow = 1 To objGrid1.iLinhasExistentes
        GridRotinas.TextMatrix(iRow, 1) = "Com Acesso"
    Next

End Sub
