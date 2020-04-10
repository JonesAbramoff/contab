VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RotinaGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotina x Grupo"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "RotinaGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7485
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   135
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RotinaGrupo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "RotinaGrupo.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "RotinaGrupo.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox SiglaRotina 
      Height          =   315
      Left            =   1305
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3915
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2670
   End
   Begin VB.ListBox Rotinas 
      Height          =   3570
      ItemData        =   "RotinaGrupo.frx":0954
      Left            =   6480
      List            =   "RotinaGrupo.frx":0956
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2715
   End
   Begin VB.TextBox Projeto 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2115
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3570
      Width           =   1650
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "RotinaGrupo.frx":0958
      Left            =   825
      List            =   "RotinaGrupo.frx":0962
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3540
      Width           =   1275
   End
   Begin VB.TextBox Classe 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3810
      MaxLength       =   50
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3585
      Width           =   1650
   End
   Begin MSFlexGridLib.MSFlexGrid GridGrupos 
      Height          =   1890
      Left            =   255
      TabIndex        =   2
      Top             =   1560
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin VB.Label Label4 
      Caption         =   "Rotinas"
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
      Left            =   6540
      TabIndex        =   9
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rotina:"
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
      Left            =   510
      TabIndex        =   8
      Top             =   870
      Width           =   645
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
      TabIndex        =   7
      Top             =   345
      Width           =   705
   End
End
Attribute VB_Name = "RotinaGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAlterado As Integer
Dim objGrid1 As AdmGrid

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

Private Function GrupoRotina_Exibe(ByVal sSiglaRotina As String) As Long
'Exibe na Tela os dados de Rotina e GrupoRotinas correspondentes à Rotina

Dim lErro As Long
Dim colGrupoRotina As New colGrupoRotina
Dim objGrupoRotina As ClassDicGrupoRotina
Dim iRow As Integer

On Error GoTo Erro_GrupoRotina_Exibe

    'Lê dados de GrupoRotinas correspondentes a esta Rotina
    lErro = GrupoRotina_Le_Rotina(sSiglaRotina, colGrupoRotina)
    If lErro Then Error 6381
    
    'Coloca sigla de rotina na Tela
    SiglaRotina.Text = sSiglaRotina
    
    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoRotina.Count
     
    'Coloca dados de GrupoRotinas no Grid
    iRow = GridGrupos.FixedRows - 1
    For Each objGrupoRotina In colGrupoRotina
        iRow = iRow + 1
        GridGrupos.TextMatrix(iRow, 0) = objGrupoRotina.sCodGrupo
        GridGrupos.TextMatrix(iRow, 1) = IIf(objGrupoRotina.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridGrupos.TextMatrix(iRow, 2) = objGrupoRotina.sProjeto
        GridGrupos.TextMatrix(iRow, 3) = objGrupoRotina.sClasse
    Next
    
    GrupoRotina_Exibe = SUCESSO
    
    Exit Function

Erro_GrupoRotina_Exibe:

    GrupoRotina_Exibe = Err

    Select Case Err
    
        Case 6381  'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174312)

    End Select

    Exit Function
    
End Function

Private Function Limpa_Tela_Local()
  
    'Limpa GridGrupos
    If objGrid1.iLinhasExistentes > 0 Then
                
        Call Grid_Limpa(objGrid1)
    
    End If

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
    objGrid1.colColuna.Add ("Projeto Customizado")
    objGrid1.colColuna.Add ("Classe Customizada")
    
   'campos de edição do grid
    objGrid1.colCampo.Add (Acesso.Name)
    objGrid1.colCampo.Add (Projeto.Name)
    objGrid1.colCampo.Add (Classe.Name)
    
    'Grid
    objGrid1.objGrid = GridGrupos
    
    'Proibe a exclusão de linhas do GRID
    objGrid1.iProibidoExcluir = 1
    
    'Proibe a inclusão de linhas no GRID
    objGrid1.iProibidoIncluir = 1
    
    'Lê os Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 6368 'Não existem Grupos
    If lErro Then Error 6369
    
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

        Case 6368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
        
        Case 6369   'Tratado na rotina chamada
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174313)

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

    Unload RotinaGrupo

End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
Dim colGrupoRotina As New colGrupoRotina
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Rotina foram informados
    If Len(SiglaRotina.Text) = 0 Then Error 6384
    
    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 6385
    
    'Armazena linhas do Grid em colGrupoRotina
    For iIndice = 1 To objGrid1.iLinhasExistentes
        colGrupoRotina.Add LOG_NAO, GridGrupos.TextMatrix(iIndice, 2), GridGrupos.TextMatrix(iIndice, 3), GridGrupos.TextMatrix(iIndice, 0), SiglaRotina.Text, IIf(GridGrupos.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO)
    Next
            
    'Grava colGrupoRotina no banco de dados (é um update)
    lErro = GrupoRotina_Grava(colGrupoRotina)
    If lErro Then Error 6386
         
    'Limpa a Tela
    Call Limpa_Tela(RotinaGrupo)
    Call Limpa_Tela_Local
  
Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 6384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ROTINA_NAO_INFORMADA", Err)
            Rotinas.SetFocus
            
        Case 6385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID_GRUPOS", Err)
    
        Case 6386  'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174314)

     End Select
        
     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
    'Limpa a Tela
    Call Limpa_Tela(RotinaGrupo)
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
Dim colModulo As New Collection
Dim vModulo As Variant
Dim sModulo As String
Dim colRotina As New Collection
Dim vRotina As Variant

On Error GoTo Erro_Rotina_Form_Load

    Me.HelpContextID = IDH_ROTINA_GRUPO
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 6374
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6375
    
    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next
    
    'Se há uma Rotina selecionada
    If Len(gsRotina) > 0 Then
        
        'Lê nome do Módulo que contém a Rotina
        lErro = Modulo_Le_Rotina(gsRotina, sModulo)
        If lErro = 6308 Then Error 6376  'Não há módulo contendo Rotina
        If lErro Then Error 6377
    
        'Seleciona sModulo na ComboBox Modulo
        Call ListBox_Select(sModulo, Modulo)
        
        'Seleciona Rotina na ListBox Rotinas
        Call ListBox_Select(gsRotina, Rotinas)
            
        'Exibe na Tela dados de Rotina e GrupoRotinas correspondentes à Rotina
        lErro = GrupoRotina_Exibe(gsRotina)
        If lErro Then Error 6379
        
        gsRotina = ""
        
    Else
        
        'Seleciona o primeiro Módulo na ComboBox Modulo
        Modulo.ListIndex = 0
     
    End If
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Rotina_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
            
        Case 6374, 6375, 6377, 6378, 6379, 6380   'Tratado na rotina chamada
        
        Case 6376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_ROTINA_INEXISTENTE", Err, gsRotina)
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174315)

    End Select
    
    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colRotina As New Collection
Dim vRotina As Variant

On Error GoTo Erro_Modulo_Click
     
    'Lê siglas de Rotinas contidas no Módulo
    lErro = Rotinas_Le_NomeModulo(Modulo.Text, colRotina)
    If lErro Then Error 6383
    
    'Limpa a ListBox Rotinas
    Rotinas.Clear
    
    'Preenche ListBox Rotinas
    For Each vRotina In colRotina
        Rotinas.AddItem (vRotina)
    Next
    
    'Limpa a Tela
    Call Limpa_Tela(RotinaGrupo)
    Call Limpa_Tela_Local
    
    Exit Sub
    
Erro_Modulo_Click:

    Select Case Err
            
        Case 6383  'Tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174316)

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

Private Sub Rotinas_DblClick()

Dim lErro As Long

On Error GoTo Erro_Rotinas_DblClick
      
    'Exibe dados de Rotina e GrupoRotinas na Tela
    If Rotinas.ListIndex > -1 Then
        
        lErro = GrupoRotina_Exibe(Rotinas.Text)
        If lErro Then Error 6382
    
    End If
                         
    Exit Sub
    
Erro_Rotinas_DblClick:

    Select Case Err
            
        Case 6382  'Tratado na rotina chamada
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174317)

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
                If lErro Then Error 6398
                
            Case 2
                Set objGridInt.objControle = Projeto
                
                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6399
                
            Case 3
                Set objGridInt.objControle = Classe
                
                'Critica da coluna 3 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6400
                
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 6401
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
    
        Case 6398, 6399, 6400, 6401
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174318)
        
    End Select

    Exit Function
    
End Function
