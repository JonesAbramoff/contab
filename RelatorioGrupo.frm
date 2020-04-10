VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RelatorioGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório x Grupo"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "RelatorioGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelatorioGrupo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "RelatorioGrupo.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "RelatorioGrupo.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboRelatorios 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3945
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   2670
   End
   Begin VB.TextBox NomeTskCustomizado 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2415
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3900
      Width           =   2145
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "RelatorioGrupo.frx":0954
      Left            =   1170
      List            =   "RelatorioGrupo.frx":095E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3870
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid GridGrupos 
      Height          =   1890
      Left            =   1260
      TabIndex        =   2
      Top             =   1365
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      GridColorFixed  =   -2147483640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Relatório:"
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
      Left            =   330
      TabIndex        =   9
      Top             =   765
      Width           =   855
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
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   705
   End
End
Attribute VB_Name = "RelatorioGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAlterado As Integer
Dim objGrid1 As AdmGrid

Private Sub ComboRelatorios_Click()

Dim lErro As Long

On Error GoTo Erro_ComboRelatorios_Click

    If ComboRelatorios.ListIndex <> -1 Then
    
        'Exibe dados dos grupos correspondentes ao relatorio selecionado
        lErro = GrupoRelatorio_Exibe(ComboRelatorios.Text)
        If lErro Then Error 32103

    End If
    
    Exit Sub
    
Erro_ComboRelatorios_Click:

    Select Case Err
    
        Case 32103 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166659)

    End Select
    
    Exit Sub

End Sub

Private Sub GridGrupos_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    

End Sub

Private Sub GridGrupos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridGrupos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridGrupos_LeaveCell()
    
    If objGrid1.iSaidaCelula = 1 Then Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridGrupos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    
End Sub

Private Sub GridGrupos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridGrupos_LostFocus()
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridGrupos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridGrupos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Private Function GrupoRelatorio_Exibe(ByVal sCodRel As String) As Long
'Exibe na Tela os dados de Relatorio e GrupoRelatorios correspondentes ao Relatorio

Dim lErro As Long
Dim colGrupoRelatorio As New Collection
Dim objGrupoRelatorio As ClassDicGrupoRelatorio
Dim iRow As Integer

On Error GoTo Erro_GrupoRelatorio_Exibe

    'Lê dados de GrupoRelatorios correspondentes a este Relatorio
    lErro = GrupoRelatorio_Le_Relatorio(sCodRel, colGrupoRelatorio)
    If lErro <> SUCESSO Then Error 32104
    
    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoRelatorio.Count
     
    'Coloca dados de GrupoRelatorios no Grid
    iRow = GridGrupos.FixedRows - 1
    For Each objGrupoRelatorio In colGrupoRelatorio
        iRow = iRow + 1
        GridGrupos.TextMatrix(iRow, 0) = objGrupoRelatorio.sCodGrupo
        GridGrupos.TextMatrix(iRow, 1) = IIf(objGrupoRelatorio.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridGrupos.TextMatrix(iRow, 2) = objGrupoRelatorio.sNomeTskCustomizado
    Next
    
    GrupoRelatorio_Exibe = SUCESSO
    
    Exit Function

Erro_GrupoRelatorio_Exibe:

    GrupoRelatorio_Exibe = Err

    Select Case Err
    
        Case 32104  'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166660)

    End Select

    Exit Function
    
End Function

Private Function Limpa_Tela_Local()
  
    'Limpa GridGrupos
    If objGrid1.iLinhasExistentes > 0 Then
                
        Call Grid_Limpa(objGrid1)
    
    End If

    ComboRelatorios.ListIndex = -1
    
End Function

Private Function Grid_Inicia() As Long

Dim lErro As Long
Dim colGrupo As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Grid_Inicia

    'Cria AdmGrid
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGrid1.colColuna.Add ("Grupo")
    objGrid1.colColuna.Add ("Tipo de Acesso")
    objGrid1.colColuna.Add ("Arquivo Customizado")
    
   'campos de edição do grid
    objGrid1.colCampo.Add (Acesso.Name)
    objGrid1.colCampo.Add (NomeTskCustomizado.Name)
    
    'Grid
    objGrid1.objGrid = GridGrupos
    
    'Proibe a exclusão de linhas do GRID
    objGrid1.iProibidoExcluir = 1
    
    'Proibe a inclusão de linhas no GRID
    objGrid1.iProibidoIncluir = 1
    
    'Lê os Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 32105 'Não existem Grupos
    If lErro Then Error 32106
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9
    
    'todas as linhas do grid
    GridGrupos.Rows = IIf(colGrupo.Count > objGrid1.iLinhasVisiveis, colGrupo.Count, objGrid1.iLinhasVisiveis) + GridGrupos.FixedRows
    
    GridGrupos.ColWidth(0) = 1200
    
    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGrid1)
    
    'Limpa os índices na primeira coluna
    For iIndice = 1 To GridGrupos.Rows - 1
    
        GridGrupos.TextMatrix(iIndice, 0) = ""
    
    Next
    
    'Alinha a primeira coluna à esquerda
    GridGrupos.ColAlignment(0) = flexAlignLeftCenter
  
    Grid_Inicia = SUCESSO

    Exit Function

Erro_Grid_Inicia:

    Grid_Inicia = Err

    Select Case Err

        Case 32105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
        
        Case 32106   'Tratado na rotina chamada
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166661)

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

    'Verifica se dados do Relatorio foram informados
    If ComboRelatorios.ListIndex = -1 Then Error 32107
    
    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 32108
    
    'Armazena linhas do Grid em colGrupoRelatorio
    For iIndice = 1 To objGrid1.iLinhasExistentes
    
        Set objGrupoRelatorio = New ClassDicGrupoRelatorio
        
        objGrupoRelatorio.iTipoDeAcesso = IIf(GridGrupos.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO)
        objGrupoRelatorio.sCodGrupo = GridGrupos.TextMatrix(iIndice, 0)
        objGrupoRelatorio.sCodRel = ComboRelatorios.Text
        objGrupoRelatorio.sNomeTskCustomizado = GridGrupos.TextMatrix(iIndice, 2)
        
        colGrupoRelatorio.Add objGrupoRelatorio
        
    Next
            
    'Grava colGrupoRelatorio no banco de dados (é um update)
    lErro = GrupoRelatorio_Grava(colGrupoRelatorio)
    If lErro Then Error 32109
         
    'Limpa a Tela
    Call Limpa_Tela(Me)
    Call Limpa_Tela_Local
  
Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 32107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_NAO_INFORMADO", Err)
            
        Case 32108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID", Err)
    
        Case 32109  'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166662)

     End Select
        
     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
    'Limpa a Tela
    Call Limpa_Tela(Me)
    Call Limpa_Tela_Local

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colModulo As New Collection
Dim vModulo As Variant
Dim sModulo As String

On Error GoTo Erro_Form_Load

    Me.HelpContextID = IDH_RELATORIO_GRUPO
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 32110
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 32111
    
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
            
        Case 32110, 32111   'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166663)

    End Select
    
    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colRelatorio As New Collection
Dim vRelatorio As Variant

On Error GoTo Erro_Modulo_Click
     
    'Lê os nomes de Relatorios associados ao Módulo
    lErro = Relatorios_Le_NomeModulo(Modulo.Text, colRelatorio)
    If lErro Then Error 32112
    
    ComboRelatorios.Clear
    
    'Preenche ComboBox de Relatorios
    For Each vRelatorio In colRelatorio
        ComboRelatorios.AddItem (vRelatorio)
    Next
    
    'Limpa a Tela
    Call Limpa_Tela(Me)
    Call Limpa_Tela_Local
    
    Exit Sub
    
Erro_Modulo_Click:

    Select Case Err
            
        Case 32112  'Tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166664)

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
                If lErro Then Error 32113
                
            Case 2
                Set objGridInt.objControle = NomeTskCustomizado
                
                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 32114
                                
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 32115
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
    
        Case 32113, 32114, 32115
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166665)
        
    End Select

    Exit Function
    
End Function
