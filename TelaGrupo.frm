VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form TelaGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tela x Grupo"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "TelaGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7500
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "TelaGrupo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "TelaGrupo.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TelaGrupo.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Classe 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4140
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3960
      Width           =   1650
   End
   Begin VB.ComboBox Acesso 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "TelaGrupo.frx":0954
      Left            =   1185
      List            =   "TelaGrupo.frx":095E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3930
      Width           =   1275
   End
   Begin VB.TextBox Projeto 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2460
      MaxLength       =   50
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3960
      Width           =   1650
   End
   Begin VB.ListBox Telas 
      Height          =   3570
      ItemData        =   "TelaGrupo.frx":097A
      Left            =   6480
      List            =   "TelaGrupo.frx":0981
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1515
      Width           =   2715
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   2670
   End
   Begin VB.TextBox Tela 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   795
      Width           =   3915
   End
   Begin MSFlexGridLib.MSFlexGrid GridGrupos 
      Height          =   1890
      Left            =   195
      TabIndex        =   2
      Top             =   1515
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
      Left            =   495
      TabIndex        =   9
      Top             =   300
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tela:"
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
      Left            =   615
      TabIndex        =   8
      Top             =   825
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Telas"
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
      Left            =   6510
      TabIndex        =   7
      Top             =   1230
      Width           =   690
   End
End
Attribute VB_Name = "TelaGrupo"
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

Private Function GrupoTela_Exibe(ByVal sTela As String) As Long
'Exibe na Tela os dados de Tela e GrupoTela correspondentes à Tela

Dim lErro As Long
Dim colGrupoTela As New colGrupoTela
Dim objGrupoTela As ClassDicGrupoTela
Dim iRow As Integer

On Error GoTo Erro_GrupoTela_Exibe

    'Lê dados de GrupoTela correspondentes a esta Tela
    lErro = GrupoTela_Le_Tela(sTela, colGrupoTela)
    If lErro Then Error 6405
    
    'Coloca nome de tela na Tela
    Tela.Text = sTela
    
    'Linhas existentes no Grid
    objGrid1.iLinhasExistentes = colGrupoTela.Count
     
    'Coloca dados de GrupoTela no Grid
    iRow = GridGrupos.FixedRows - 1
    For Each objGrupoTela In colGrupoTela
        iRow = iRow + 1
        GridGrupos.TextMatrix(iRow, 0) = objGrupoTela.sCodGrupo
        GridGrupos.TextMatrix(iRow, 1) = IIf(objGrupoTela.iTipoDeAcesso = COM_ACESSO, "Com Acesso", "Sem Acesso")
        GridGrupos.TextMatrix(iRow, 2) = objGrupoTela.sProjeto
        GridGrupos.TextMatrix(iRow, 3) = objGrupoTela.sClasse
    Next
    
    GrupoTela_Exibe = SUCESSO
    
    Exit Function

Erro_GrupoTela_Exibe:

    GrupoTela_Exibe = Err

    Select Case Err
    
        Case 6405  'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174637)

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
    If lErro = 6365 Then Error 6406 'Não existem Grupos
    If lErro Then Error 6407
    
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

        Case 6406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
        
        Case 6407   'Tratado na rotina chamada
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174638)

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

    Unload TelaGrupo

End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
Dim colGrupoTela As New colGrupoTela
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Tela foram informados
    If Len(Tela.Text) = 0 Then Error 6408
    
    'Verifica se Grid está preenchido
    If objGrid1.iLinhasExistentes <= 0 Then Error 6409
    
    'Armazena linhas do Grid em colGrupoTela
    For iIndice = 1 To objGrid1.iLinhasExistentes
        colGrupoTela.Add GridGrupos.TextMatrix(iIndice, 2), GridGrupos.TextMatrix(iIndice, 3), GridGrupos.TextMatrix(iIndice, 0), IIf(GridGrupos.TextMatrix(iIndice, 1) = "Com Acesso", COM_ACESSO, SEM_ACESSO), Tela.Text
    Next
            
    'Grava colGrupoTela no banco de dados (é um update)
    lErro = GrupoTela_Grava(colGrupoTela)
    If lErro Then Error 6410
         
    'Limpa a Tela
    Call Limpa_Tela(TelaGrupo)
    Call Limpa_Tela_Local
  
Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 6408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_INFORMADA", Err)
            Telas.SetFocus
            
        Case 6409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DADOS_GRID_GRUPOS", Err)
    
        Case 6410  'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174639)

     End Select
        
     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
    'Limpa a Tela
    Call Limpa_Tela(TelaGrupo)
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
Dim colTela As New Collection
Dim vTela As Variant

On Error GoTo Erro_Tela_Form_Load

    Me.HelpContextID = IDH_TELA_GRUPO
    
    'Inicializa Grid
    lErro = Grid_Inicia()
    If lErro Then Error 6411
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6412
    
    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next
    
    'Se há uma Tela selecionada
    If Len(gsTela) > 0 Then
        
        'Lê nome do Módulo que contém a Tela
        lErro = Modulo_Le_Tela(gsTela, sModulo)
        If lErro = 6342 Then Error 6413  'Não há módulo contendo Tela
        If lErro Then Error 6414
    
        'Seleciona sModulo na ComboBox Modulo
        Call ListBox_Select(sModulo, Modulo)
        
        'Seleciona Tela na ListBox Telas
        Call ListBox_Select(gsTela, Telas)
            
        'Exibe na Tela dados de Tela e GrupoTela correspondentes à Tela
        lErro = GrupoTela_Exibe(gsTela)
        If lErro Then Error 6415
        
        gsTela = ""
        
    Else
        
        'Seleciona o primeiro Módulo na ComboBox Modulo
        Modulo.ListIndex = 0
     
    End If
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Tela_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
            
        Case 6411, 6412, 6414, 6378, 6415, 6380   'Tratado na rotina chamada
        
        Case 6413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_TELA_INEXISTENTE", Err, gsTela)
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174640)

    End Select
    
    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colTela As New Collection
Dim vTela As Variant

On Error GoTo Erro_Modulo_Click
     
    'Lê siglas de Telas contidas no Módulo
    lErro = Telas_Le_NomeModulo(Modulo.Text, colTela)
    If lErro Then Error 6416
    
    'Limpa a ListBox Telas
    Telas.Clear
    
    'Preenche ListBox Telas
    For Each vTela In colTela
        Telas.AddItem (vTela)
    Next
    
    'Limpa a Tela
    Call Limpa_Tela(TelaGrupo)
    Call Limpa_Tela_Local
    
    Exit Sub
    
Erro_Modulo_Click:

    Select Case Err
            
        Case 6416  'Tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174641)

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
                If lErro Then Error 6417
                
            Case 2
                Set objGridInt.objControle = Projeto
                
                'critica da coluna 2 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6418
                
            Case 3
                Set objGridInt.objControle = Classe
                
                'Critica da coluna 3 vazia
                lErro = Grid_Abandona_Celula(objGridInt)
                If lErro Then Error 6419
                
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 6420
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
    
        Case 6417, 6418, 6419, 6420
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174642)
        
    End Select

    Exit Function
    
End Function

Private Sub Telas_DblClick()

Dim lErro As Long

On Error GoTo Erro_Telas_DblClick
      
    'Exibe dados de Tela e GrupoTela na Tela
    If Telas.ListIndex > -1 Then
        
        lErro = GrupoTela_Exibe(Telas.Text)
        If lErro Then Error 6439
    
    End If
                         
    Exit Sub
    
Erro_Telas_DblClick:

    Select Case Err
            
        Case 6439  'Tratado na rotina chamada
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174643)

    End Select
    
    Exit Sub

End Sub
