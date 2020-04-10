VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GrupoRelatorio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo x Relatório"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "GrupoRelatorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6510
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GrupoRelatorio.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "GrupoRelatorio.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "GrupoRelatorio.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
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
      Left            =   1275
      Picture         =   "GrupoRelatorio.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
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
      Left            =   3480
      Picture         =   "GrupoRelatorio.frx":196E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4380
      Width           =   1830
   End
   Begin VB.ComboBox Grupo 
      Height          =   315
      Left            =   855
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   1245
   End
   Begin VB.TextBox NomeTskCustomizado 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4485
      MaxLength       =   64
      TabIndex        =   4
      Top             =   3285
      Width           =   2445
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "GrupoRelatorio.frx":2B50
      Left            =   3195
      List            =   "GrupoRelatorio.frx":2B5A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3240
      Width           =   1275
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   3435
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   2670
   End
   Begin MSFlexGridLib.MSFlexGrid GridRelatorios 
      Height          =   1890
      Left            =   105
      TabIndex        =   2
      Top             =   825
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   3334
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
      Left            =   120
      TabIndex        =   6
      Top             =   375
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
      Left            =   2610
      TabIndex        =   5
      Top             =   375
      Width           =   705
   End
End
Attribute VB_Name = "GrupoRelatorio"
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
        GridRelatorios.TextMatrix(iRow, 1) = "Sem Acesso"
    Next

End Sub

Private Sub GridRelatorios_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If


End Sub

Private Sub GridRelatorios_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridRelatorios_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridRelatorios_LeaveCell()

    If objGrid1.iSaidaCelula = 1 Then Call Saida_Celula(objGrid1)

End Sub

Private Sub GridRelatorios_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Private Sub GridRelatorios_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridRelatorios_LostFocus()

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridRelatorios_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridRelatorios_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Private Function GrupoRelatorio_Exibe(ByVal sGrupo As String, ByVal sModulo As String) As Long
'Exibe na Tela os dados de GrupoRelatorios correspondentes ao Grupo e ao Módulo (relatorios dentro do módulo)

Dim lErro As Long
Dim colGrupoRelatorio As New Collection
Dim objGrupoRelatorio As ClassDicGrupoRelatorio
Dim iRow As Integer, iNovoNumRows As Integer

On Error GoTo Erro_GrupoRelatorio_Exibe

    'Lê dados de GrupoRelatorios correspondentes ao Grupo e ao Módulo
    lErro = GrupoRelatorio_Le_GrupoModulo(sGrupo, sModulo, colGrupoRelatorio)
    If lErro Then Error 32090

    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoRelatorio.Count

    iNovoNumRows = IIf(colGrupoRelatorio.Count > objGrid1.iLinhasVisiveis, colGrupoRelatorio.Count, objGrid1.iLinhasVisiveis) + 1

    If iNovoNumRows <> GridRelatorios.Rows Then
    
        GridRelatorios.Rows = iNovoNumRows
        
        're-inicializa o grid
        Call Grid_Inicializa(objGrid1)
    
    End If
    
    'Coloca dados de GrupoRelatorios no Grid
    iRow = GridRelatorios.FixedRows - 1
    For Each objGrupoRelatorio In colGrupoRelatorio
        iRow = iRow + 1
        GridRelatorios.TextMatrix(iRow, 0) = objGrupoRelatorio.sCodRel
        GridRelatorios.TextMatrix(iRow, 1) = IIf(objGrupoRelatorio.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridRelatorios.TextMatrix(iRow, 2) = objGrupoRelatorio.sNomeTskCustomizado
    Next

    GrupoRelatorio_Exibe = SUCESSO

    Exit Function

Erro_GrupoRelatorio_Exibe:

    GrupoRelatorio_Exibe = Err

    Select Case Err

        Case 32090 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161723)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_Local()

Dim iLinha As Integer
    
    'Limpa ComboBox Grupo
    Grupo.ListIndex = -1
    
    'Limpa GridRelatorios
    If objGrid1.iLinhasExistentes > 0 Then
        
        Call Grid_Limpa(objGrid1)
        
        'Limpa a primeira coluna
        For iLinha = 1 To GridRelatorios.Rows - 1
            GridRelatorios.TextMatrix(iLinha, 0) = ""
        Next

    End If

End Function

Private Function Grid_Inicia() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Grid_Inicia

    'Cria AdmGrid
    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    'titulos do grid
    objGrid1.colColuna.Add ("Relatório")
    objGrid1.colColuna.Add ("Tipo de Acesso")
    objGrid1.colColuna.Add ("Arquivo Customizado")

   'campos de edição do grid
    objGrid1.colCampo.Add (Acesso.Name)
    objGrid1.colCampo.Add (NomeTskCustomizado.Name)

    'Grid
    objGrid1.objGrid = GridRelatorios

    'Proibe a exclusão de linhas do GRID
    objGrid1.iProibidoExcluir = 1

    'Proibe a inclusão de linhas no GRID
    objGrid1.iProibidoIncluir = 1

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9

    'todas as linhas do grid
    GridRelatorios.Rows = objGrid1.iLinhasVisiveis + GridRelatorios.FixedRows

    GridRelatorios.ColWidth(0) = 3900

    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGrid1)

    'Limpa os índices na primeira coluna
    For iIndice = 1 To GridRelatorios.Rows - 1

        GridRelatorios.TextMatrix(iIndice, 0) = ""

    Next

    'Alinha a primeira coluna à esquerda
    GridRelatorios.ColAlignment(0) = flexAlignLeftCenter

    Grid_Inicia = SUCESSO

    Exit Function

Erro_Grid_Inicia:

    Grid_Inicia = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161724)

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

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim colGrupoRelatorio As New Collection
Dim objGrupoRelatorio As ClassDicGrupoRelatorio
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados obrigatorios foram informados
    If Len(Grupo.Text) = 0 Then Error 32091

    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 32092

    'Armazena linhas do Grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
    
        Set objGrupoRelatorio = New ClassDicGrupoRelatorio
        
        objGrupoRelatorio.iTipoDeAcesso = IIf(GridRelatorios.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO)
        objGrupoRelatorio.sCodGrupo = Grupo.Text
        objGrupoRelatorio.sCodRel = GridRelatorios.TextMatrix(iIndice, 0)
        objGrupoRelatorio.sNomeTskCustomizado = GridRelatorios.TextMatrix(iIndice, 2)
        
        colGrupoRelatorio.Add objGrupoRelatorio
        
    Next

    'Grava colGrupoRelatorio no banco de dados (é um update)
    lErro = GrupoRelatorio_Grava2(Modulo.Text, colGrupoRelatorio)
    If lErro Then Error 32093

    'Limpa a Tela
    Call Limpa_Tela_Local

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 32091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_INFORMADO", Err)
            Grupo.SetFocus
            
        Case 32092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID", Err)

        Case 32093  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161725)

     End Select

     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    'Limpa a Tela
    Call Limpa_Tela_Local

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colGrupo As New Collection
Dim vGrupo As Variant
Dim colModulo As New Collection
Dim vModulo As Variant

On Error GoTo Erro_Form_Load

    Me.HelpContextID = IDH_GRUPO_RELATORIO
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 32094
    
    'Lê Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 32095
    If lErro Then Error 32096

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
    If lErro Then Error 32097

    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next

    'Seleciona o primeiro Módulo na ComboBox Modulo
    Modulo.ListIndex = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err

        Case 32094, 32096, 32097  'Tratado na rotina chamada
        
        Case 32095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161726)

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
            For iLinha = 1 To GridRelatorios.Rows - 1
                GridRelatorios.TextMatrix(iLinha, 0) = ""
            Next

        End If

        'Preenche Grid com dados de GrupoRelatorios associados ao Grupo e aos relatorios do Módulo
        lErro = GrupoRelatorio_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then Error 32098
        
    End If

    Exit Sub

Erro_Modulo_Click:

    Select Case Err

        Case 32098  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161727)

    End Select

    Exit Sub

End Sub

Private Sub NomeTskCustomizado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub NomeTskCustomizado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub NomeTskCustomizado_LostFocus()

    Set objGrid1.objControle = NomeTskCustomizado
    Call Grid_Campo_Libera_Foco(objGrid1)

End Sub

Private Sub Grupo_Click()

Dim lErro As Long

On Error GoTo Erro_Grupo_Click

    If Len(Grupo.Text) > 0 And Len(Modulo.Text) > 0 Then
    
        'Preenche Grid com dados de GrupoRelatorios associados ao Grupo e aos relatorios do Módulo
        lErro = GrupoRelatorio_Exibe(Grupo.Text, Modulo.Text)
        If lErro Then Error 32099
        
    End If
    
    Exit Sub

Erro_Grupo_Click:

    Select Case Err

        Case 32099  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161728)

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
                If lErro Then Error 32100

            Case 2
                Set objGridInt.objControle = NomeTskCustomizado

                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 32101

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 32102

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 32100, 32101, 32102
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161729)

    End Select

    Exit Function

End Function


Private Sub SelecionaTodas_Click()

Dim iRow As Integer

    For iRow = 1 To objGrid1.iLinhasExistentes
        GridRelatorios.TextMatrix(iRow, 1) = "Com Acesso"
    Next

End Sub
