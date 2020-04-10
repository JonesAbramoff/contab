VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GrupoTela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo x Tela"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   9480
   Icon            =   "GrupoTela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox NomeTela 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   50
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3420
      Width           =   2475
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7650
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   105
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "GrupoTela.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "GrupoTela.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GrupoTela.frx":07FA
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
      Height          =   675
      Left            =   1650
      Picture         =   "GrupoTela.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4605
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
      Height          =   675
      Left            =   3825
      Picture         =   "GrupoTela.frx":196E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4605
      Width           =   1830
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
   Begin VB.TextBox Projeto 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4770
      MaxLength       =   50
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3345
      Width           =   1515
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "GrupoTela.frx":2B50
      Left            =   3450
      List            =   "GrupoTela.frx":2B5A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3300
      Width           =   1275
   End
   Begin VB.TextBox Classe 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6375
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3330
      Width           =   1530
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
   Begin MSFlexGridLib.MSFlexGrid GridTelas 
      Height          =   1635
      Left            =   345
      TabIndex        =   2
      Top             =   900
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   2884
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
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
      TabIndex        =   7
      Top             =   435
      Width           =   615
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
      TabIndex        =   6
      Top             =   435
      Width           =   705
   End
End
Attribute VB_Name = "GrupoTela"
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
        GridTelas.TextMatrix(iRow, 1) = "Sem Acesso"
    Next

End Sub

Private Sub GridTelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If


End Sub

Private Sub GridTelas_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridTelas_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridTelas_LeaveCell()

    If objGrid1.iSaidaCelula = 1 Then Call Saida_Celula(objGrid1)

End Sub

Private Sub GridTelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Private Sub GridTelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridTelas_LostFocus()

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridTelas_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridTelas_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Private Function GrupoTela_Exibe(ByVal sGrupo As String, ByVal sModulo As String) As Long
'Exibe na Tela os dados de GrupoTela correspondentes ao Grupo e ao Módulo (telas dentro do módulo)

Dim lErro As Long
Dim colGrupoTela As New colGrupoTela
Dim objGrupoTela As ClassDicGrupoTela
Dim iRow As Integer, iNovoNumRows As Integer
Dim objTela As New ClassDicTela

On Error GoTo Erro_GrupoTela_Exibe

    'Lê dados de GrupoTela correspondentes ao Grupo e ao Módulo
    lErro = GrupoTela_Le_GrupoModulo(sGrupo, sModulo, colGrupoTela)
    If lErro Then gError 6477

    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoTela.Count

    iNovoNumRows = IIf(colGrupoTela.Count > objGrid1.iLinhasVisiveis, colGrupoTela.Count, objGrid1.iLinhasVisiveis) + 1

    If iNovoNumRows <> GridTelas.Rows Then
    
        GridTelas.Rows = iNovoNumRows
        
        're-inicializa o grid
        Call Grid_Inicializa(objGrid1)
    
    End If
    
    'Coloca dados de GrupoTela no Grid
    iRow = GridTelas.FixedRows - 1
    For Each objGrupoTela In colGrupoTela
        iRow = iRow + 1
        
        objTela.sNome = objGrupoTela.sNomeTela
        
        lErro = Tela_Le(objTela)
        If lErro Then gError 82261
        
        If Len(Trim(objTela.sDescricao)) = 0 Then
            GridTelas.TextMatrix(iRow, 0) = objGrupoTela.sNomeTela
        Else
            GridTelas.TextMatrix(iRow, 0) = objTela.sDescricao
        End If
        GridTelas.TextMatrix(iRow, 1) = IIf(objGrupoTela.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridTelas.TextMatrix(iRow, 2) = objGrupoTela.sProjeto
        GridTelas.TextMatrix(iRow, 3) = objGrupoTela.sClasse
        GridTelas.TextMatrix(iRow, 4) = objGrupoTela.sNomeTela
    Next

    GrupoTela_Exibe = SUCESSO

    Exit Function

Erro_GrupoTela_Exibe:

    GrupoTela_Exibe = gErr

    Select Case gErr

        Case 6477, 82261 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161737)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_Local()

Dim iLinha As Integer
    
    'Limpa ComboBox Grupo
    Grupo.ListIndex = -1
    
    'Limpa GridTelas
    If objGrid1.iLinhasExistentes > 0 Then
        
        Call Grid_Limpa(objGrid1)
        
        'Limpa a primeira coluna
        For iLinha = 1 To GridTelas.Rows - 1
            GridTelas.TextMatrix(iLinha, 0) = ""
        Next

    End If

End Function

Private Function Grid_Inicia() As Long

Dim lErro As Long
Dim colTela As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Grid_Inicia

    'Cria AdmGrid
    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    'titulos do grid
    
    objGrid1.colColuna.Add ("Descrição")
    objGrid1.colColuna.Add ("Tipo de Acesso")
    objGrid1.colColuna.Add ("Projeto Customizado")
    objGrid1.colColuna.Add ("Classe Customizada")
    objGrid1.colColuna.Add ("Tela")
    
   'campos de edição do grid
    objGrid1.colCampo.Add (Acesso.Name)
    objGrid1.colCampo.Add (Projeto.Name)
    objGrid1.colCampo.Add (Classe.Name)
    objGrid1.colCampo.Add (NomeTela.Name)

    'Grid
    objGrid1.objGrid = GridTelas

    'Proibe a exclusão de linhas do GRID
    objGrid1.iProibidoExcluir = 1

    'Proibe a inclusão de linhas no GRID
    objGrid1.iProibidoIncluir = 1

    'Lê os Grupos existentes no BD
    lErro = Telas_Le_Todas(colTela)
    If lErro Then Error 6478

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9

    'todas as linhas do grid
    GridTelas.Rows = IIf(colTela.Count > objGrid1.iLinhasVisiveis, colTela.Count, objGrid1.iLinhasVisiveis) + GridTelas.FixedRows

    GridTelas.ColWidth(0) = 4100

    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid1)

    'Limpa os índices na primeira coluna
    For iIndice = 1 To GridTelas.Rows - 1

        GridTelas.TextMatrix(iIndice, 0) = ""

    Next

    'Alinha a primeira coluna à esquerda
    GridTelas.ColAlignment(0) = flexAlignLeftCenter

    Grid_Inicia = SUCESSO

    Exit Function

Erro_Grid_Inicia:

    Grid_Inicia = Err

    Select Case Err

        Case 6478   'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161738)

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

    Unload GrupoTela

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim colGrupoTela As New colGrupoTela
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Tela foram informados
    If Len(Grupo.Text) = 0 Then Error 6479

    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 6480

    'Armazena linhas do Grid em colGrupoTela
    For iIndice = 1 To objGrid1.iLinhasExistentes
        colGrupoTela.Add GridTelas.TextMatrix(iIndice, 2), GridTelas.TextMatrix(iIndice, 3), Grupo.Text, IIf(GridTelas.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO), GridTelas.TextMatrix(iIndice, 4)
    Next

    'Grava colGrupoTela no banco de dados (é um update)
    lErro = GrupoTela_Grava2(Modulo.Text, colGrupoTela)
    If lErro Then Error 6481

    'Limpa a Tela
    Call Limpa_Tela_Local

Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 6479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_INFORMADO", Err)
            Grupo.SetFocus

        Case 6480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID_TELAS", Err)

        Case 6481  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161739)

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

On Error GoTo Erro_GrupoTela_Form_Load

    Me.HelpContextID = IDH_GRUPO_TELA
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 6482
    
    'Lê Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 6483
    If lErro Then Error 6484

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
    If lErro Then Error 6485

    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next

    'Seleciona o primeiro Módulo na ComboBox Modulo
    Modulo.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_GrupoTela_Form_Load:

    lErro_Chama_Tela = Err
    Select Case Err

        Case 6482, 6484, 6485  'Tratado na rotina chamada
        
        Case 6483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161740)

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
            For iLinha = 1 To GridTelas.Rows - 1
                GridTelas.TextMatrix(iLinha, 0) = ""
            Next

        End If

        'Preenche Grid com dados de GrupoTela associados ao Grupo e às telas do Módulo
        lErro = GrupoTela_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then gError 6486
        
    End If

    Exit Sub

Erro_Modulo_Click:

    Select Case gErr

        Case 6486  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161741)

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
    
        'Preenche Grid com dados de GrupoTela associados ao Grupo e às telas do Módulo
        lErro = GrupoTela_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then Error 6487
        
    End If
    
    Exit Sub

Erro_Grupo_Click:

    Select Case Err

        Case 6487  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161742)

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
                If lErro Then Error 6488

            Case 2
                Set objGridInt.objControle = Projeto

                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6489

            Case 3
                Set objGridInt.objControle = Classe

                'Critica da coluna 3 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6490

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 6491

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 6488, 6489, 6490, 6491
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161743)

    End Select

    Exit Function

End Function

Private Sub SelecionaTodas_Click()

Dim iRow As Integer

    For iRow = 1 To objGrid1.iLinhasExistentes
        GridTelas.TextMatrix(iRow, 1) = "Com Acesso"
    Next

End Sub
