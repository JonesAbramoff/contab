VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl SegmentosMAT 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   8430
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6585
      ScaleHeight     =   495
      ScaleWidth      =   1650
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   1710
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "SegmentosMAT.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "SegmentosMAT.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "SegmentosMAT.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Tipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SegmentosMAT.ctx":080A
      Left            =   1140
      List            =   "SegmentosMAT.ctx":080C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1020
      Width           =   1400
   End
   Begin VB.ComboBox Delimitador 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SegmentosMAT.ctx":080E
      Left            =   3480
      List            =   "SegmentosMAT.ctx":081B
      TabIndex        =   3
      Top             =   1020
      Width           =   1065
   End
   Begin VB.ComboBox Preenchimento 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SegmentosMAT.ctx":0828
      Left            =   4560
      List            =   "SegmentosMAT.ctx":082A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   3705
   End
   Begin VB.ComboBox Formato 
      Height          =   315
      ItemData        =   "SegmentosMAT.ctx":082C
      Left            =   1335
      List            =   "SegmentosMAT.ctx":082E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2500
   End
   Begin MSMask.MaskEdBox Tamanho 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   1020
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridSegmentos 
      Height          =   2580
      Left            =   135
      TabIndex        =   5
      Top             =   900
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   4551
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Segmentos"
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
      Left            =   165
      TabIndex        =   11
      Top             =   675
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formato de:"
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
      Left            =   210
      TabIndex        =   10
      Top             =   270
      Width           =   1020
   End
End
Attribute VB_Name = "SegmentosMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'coluna referente ao tipo no grid da tela segmentos
Const COL_TIPO = 1
'coluna referente ao tamanho no grid da tela segmentos
Const COL_TAMANHO = 2
'coluna referente ao delimitador no grid da tela segmentos
Const COL_DELIMITADOR = 3
'coluna referente ao preenchimento no grid da tela segmentos
Const COL_PREENCHIMENTO = 4

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Dim sCodigo As String

Function Trata_Parametros(Optional objsegmento As ClassSegmento) As Long
     
     iAlterado = 0
     
    Trata_Parametros = SUCESSO
     
End Function


Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 44885

    Call Grid_Limpa(objGrid1)
    
    Formato.ListIndex = -1
    
    GridSegmentos.Enabled = False
    
    iAlterado = 0

    Call Reset_Fest
    Call Reset_Contab
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err

        Case 44885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174378)

    End Select

    Exit Sub
    
End Sub

Function Gravar_Registro() As Long

Dim iTamanho As Integer
Dim iTotalTamanho As Integer
Dim iLinha As Integer
Dim lErro As Long
Dim colSegmentos As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
        
    'Verifica se pelo menos uma linha do Grid est� preenchida
    If objGrid1.iLinhasExistentes = 0 Then gError 44886
    
    iTotalTamanho = 0
    
    'percorre as linhas da coluna tamanho
    For iLinha = 1 To objGrid1.iLinhasExistentes
        
        'verifica se nao foi preenchido o tamanho
        If Len(Trim(GridSegmentos.TextMatrix(iLinha, COL_TAMANHO))) = 0 Then Error 44888
        'soma o valor total da coluna tamanho no grid
        iTotalTamanho = iTotalTamanho + CInt(GridSegmentos.TextMatrix(iLinha, COL_TAMANHO))
    
    Next
                  
    'verifica se tamanho conta ultrapassou tamanho pre_definido
    If sCodigo = SEGMENTO_PRODUTO And iTotalTamanho > STRING_PRODUTO Then
        gError 44884
    'verifica se tamanho ccl ultrapassou tamanho pre_definido
    ElseIf sCodigo = SEGMENTO_CCL And iTotalTamanho > STRING_CCL Then
        gError 44887
    '########################
    'Inserido por Wagner
    ElseIf sCodigo = SEGMENTO_NATMOVCTA And iTotalTamanho > STRING_NATMOVCTA_CODIGO Then
        gError 136640
    '########################
    End If

    'Preenche a colSegmentos com as informacoes contidas no Grid
    lErro = Grid_Segmentos(colSegmentos)
    If lErro <> SUCESSO Then gError 44889

    If sCodigo = SEGMENTO_PRODUTO Then
    
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_Produto", colSegmentos)
        If lErro <> SUCESSO Then gError 44890
        
    ElseIf sCodigo = SEGMENTO_CCL Then
        
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_Ccl", colSegmentos)
        If lErro <> SUCESSO Then gError 44891
    
    '################################
    'Inserido por Wagner
    ElseIf sCodigo = SEGMENTO_NATMOVCTA Then
        
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_NatMovCta", colSegmentos)
        If lErro <> SUCESSO Then gError 136641
    '################################
    End If
        
    GL_objMDIForm.MousePointer = vbDefault
        
    Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 44884
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_PRODUTO_MAIOR_PERMITIDO", gErr, iTotalTamanho, STRING_PRODUTO)
        
        Case 44886
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_DE_DADOS", gErr)

        Case 44887
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_CCL_MAIOR_PERMITIDO", gErr, iTotalTamanho, STRING_CCL)
        
        Case 44888
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", gErr)
            
        Case 44889, 44890, 44891, 136641 'Inserido por Wagner
        
        '##############################
        'Inserido por Wagner
        Case 136640
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_NATMOVCTA_MAIOR_PERMITIDO", gErr, iTotalTamanho, STRING_NATMOVCTA_CODIGO)
        '##############################
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174379)

    End Select

    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 44892

    Call Grid_Limpa(objGrid1)
    
    Formato.ListIndex = -1
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 44892 'cancelou operacao de gravacao , continua execucao normal
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174380)
        
    End Select

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid1 = Nothing
    
End Sub

Private Sub Formato_Click()

Dim lErro As Long
Dim objsegmento As New ClassSegmento
Dim colSegmento As New Collection
Dim iPossui_Produto As Integer
Dim iPossui_Ccl As Integer
Dim iPossui_NatMovCta As Integer 'Inserido por Wagner

On Error GoTo Erro_Formato_Click

    If Formato.ListIndex = -1 Then Exit Sub

    'Situacao qdo usuario ja cancelou troca de formato com o grid alterado
    If iAlterado = REGISTRO_CANCELADO Then
        iAlterado = REGISTRO_ALTERADO
        Exit Sub
    End If
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 44893

    'atualiza o sCodigo com o Formato corrente
    sCodigo = gobjColCodigoSegmento.Codigo(Formato.Text)

    If sCodigo = SEGMENTO_PRODUTO Then

        'faz verificacao se ja existe pelo menos um produto cadastrado na tabela PlanoConta no BD
        lErro = CF("Produto_ExisteProduto", iPossui_Produto)
        If lErro <> SUCESSO Then Error 44894
    
        'Produto ja est� cadastrado
        If iPossui_Produto = POSSUI_PRODUTO Then

            'desabilita a edicao dos campos Tipo e Tamanho do Grid
            Tipo.Enabled = False
            Tamanho.Enabled = False

            'desabilita a inclusao e exclusao de segmentos no Grid
            objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
            objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
            
        Else
            
            Tipo.Enabled = True
            Tamanho.Enabled = True

            objGrid1.iProibidoExcluir = 0
            objGrid1.iProibidoIncluir = 0

        End If
        
    ElseIf sCodigo = SEGMENTO_CCL Then
    
        'faz verificacao se ja existe pelo menos um centro de custo cadastrado no BD
        lErro = CF("Ccl_ExisteCcl", iPossui_Ccl)
        If lErro <> SUCESSO Then Error 44895
    
        'Centro de Custo j� cadastrado
        If iPossui_Ccl = POSSUI_CCL Then

            'desabilita a edicao dos campos Tipo e Tamanho do Grid
            Tipo.Enabled = False
            Tamanho.Enabled = False

            'desabilita a inclusao e exclusao de segmentos no Grid
            objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
            objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
            
            
        Else
            
            Tipo.Enabled = True
            Tamanho.Enabled = True

            objGrid1.iProibidoExcluir = 0
            objGrid1.iProibidoIncluir = 0
            
        End If
    
    '#############################################
    'Inserido por Wagner
    ElseIf sCodigo = SEGMENTO_NATMOVCTA Then
    
        'faz verificacao se ja existe pelo menos um centro de custo cadastrado no BD
        lErro = CF("NatMovCta_ExisteNatMovCta", iPossui_NatMovCta)
        If lErro <> SUCESSO Then Error 44895
    
        'Centro de Custo j� cadastrado
        If iPossui_NatMovCta = POSSUI_NATMOVCTA Then

            'desabilita a edicao dos campos Tipo e Tamanho do Grid
            Tipo.Enabled = False
            Tamanho.Enabled = False

            'desabilita a inclusao e exclusao de segmentos no Grid
            objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
            objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
            
        Else
            
            Tipo.Enabled = True
            Tamanho.Enabled = True

            objGrid1.iProibidoExcluir = 0
            objGrid1.iProibidoIncluir = 0
            
        End If
    '#############################################
    End If
    
    'preenche o obj com o formato corrente para usar em Segmento_Le_Codigo
    objsegmento.sCodigo = gobjColCodigoSegmento.Codigo(Formato.Text)

    'preenche toda colecao(colSegmento) em relacao ao formato corrente
    lErro = CF("Segmento_Le_Codigo", objsegmento, colSegmento)
    If lErro <> SUCESSO Then Error 44896

    Call Grid_Limpa(objGrid1)
    
    objGrid1.iLinhasExistentes = 0

    'preenche todo o grid da tabela segmento
    For Each objsegmento In colSegmento

        'coloca o tipo no grid da tela
        GridSegmentos.TextMatrix(objsegmento.iNivel, COL_TIPO) = gobjColTipoSegmento.Descricao(objsegmento.iTipo)

        'coloca o tamanho no grid da tela
        GridSegmentos.TextMatrix(objsegmento.iNivel, COL_TAMANHO) = objsegmento.iTamanho

        'coloca os delimitadores no grid da tela
        GridSegmentos.TextMatrix(objsegmento.iNivel, COL_DELIMITADOR) = objsegmento.sDelimitador

        'coloca o preenchimento no grid da tela
        GridSegmentos.TextMatrix(objsegmento.iNivel, COL_PREENCHIMENTO) = gobjColPreenchimento.Descricao(objsegmento.iPreenchimento)

        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1

    Next

    iAlterado = 0
    
    GridSegmentos.Enabled = True

    Exit Sub

Erro_Formato_Click:

    Select Case Err

        Case 44893
                Formato.Text = gobjColCodigoSegmento.Descricao(sCodigo)

        Case 44894, 44895, 44896

        Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174381)

    End Select

End Sub

Private Sub Tipo_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Tamanho_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Private Sub Tamanho_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Tamanho_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid1.objControle = Tamanho
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Delimitador_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Delimitador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Delimitador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Delimitador
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Preenchimento_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Preenchimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Preenchimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Preenchimento
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridSegmentos_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    

End Sub

Private Sub GridSegmentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridSegmentos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridSegmentos_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridSegmentos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    
End Sub

Private Sub GridSegmentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSegmentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridSegmentos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridSegmentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sDescricao As String 'guarda o formato
Dim iPossui_Conta As Integer 'verifica a existencia de conta em PlanoConta no BD
 
    'inicializa sCodigo com produto , ele so se altera em Formato_Click()
    sCodigo = SEGMENTO_PRODUTO
          
    Set objGrid1 = New AdmGrid
           
    'inicializacao do grid
    Call Inicializa_Grid_Segmento
    
    'inicializar os formatos
    For iIndice = 1 To gobjColCodigoSegmento.Count
    
        '####################################
        'Alterado por Wagner
        If gobjColCodigoSegmento.Item(iIndice).sCodigo = SEGMENTO_PRODUTO Or gobjColCodigoSegmento.Item(iIndice).sCodigo = SEGMENTO_CCL Or gobjColCodigoSegmento.Item(iIndice).sCodigo = SEGMENTO_NATMOVCTA Then
            Formato.AddItem gobjColCodigoSegmento.Item(iIndice).sDescricao
        End If
        '####################################
    Next
                 
    'inicializar os tipos
    For iIndice = 1 To gobjColTipoSegmento.Count
        Tipo.AddItem gobjColTipoSegmento.Item(iIndice).sDescricao
    Next

    'inicializar os preenchimentos
    For iIndice = 1 To gobjColPreenchimento.Count
        Preenchimento.AddItem gobjColPreenchimento.Item(iIndice).sDescricao
    Next

    'coloca a descricao referente ao produto em sDescricao
    sDescricao = gobjColCodigoSegmento.Descricao(SEGMENTO_PRODUTO)

    'mostra o formato produto como formato inicial
    For iIndice = 0 To Formato.ListCount - 1
        If Formato.List(iIndice) = sDescricao Then
            Formato.ListIndex = iIndice
            Exit For
        End If
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

   lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case COL_TIPO
                
                lErro = Saida_Celula_Tipo(objGridInt)
                If lErro <> SUCESSO Then Error 44896

            Case COL_TAMANHO
                
                lErro = Saida_Celula_Tamanho(objGridInt)
                If lErro <> SUCESSO Then Error 44897

            Case COL_DELIMITADOR
            
                lErro = Saida_Celula_Delimitador(objGridInt)
                If lErro <> SUCESSO Then Error 44898
                
                
             Case COL_PREENCHIMENTO
             
                lErro = Saida_Celula_Preenchimento(objGridInt)
                If lErro <> SUCESSO Then Error 44899
                   

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44900

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 44896, 44897, 44898, 44899
        
        Case 44900
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174382)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Delimitador(objGridInt As AdmGrid) As Long
'faz a critica da celula delimitador do grid que est� deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Delimitador

    Set objGridInt.objControle = Delimitador
    
    Delimitador.Text = Trim(Delimitador.Text)
    
    If Len(Delimitador.Text) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    If Len(Trim(Delimitador.Text)) > 1 Then Error 44897
    
    If Delimitador.Text = SEPARADOR Then Error 44898
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44899

    Saida_Celula_Delimitador = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Delimitador:

    Saida_Celula_Delimitador = Err
    
    Select Case Err
    
        Case 44897
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                 
        Case 44898
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 44899
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174383)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tamanho(objGridInt As AdmGrid) As Long
'faz a critica da celula Tamanho do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tamanho

    Set objGridInt.objControle = Tamanho
    
    'verifica se foi preenchido o tamanho
    If Len(Trim(Tamanho.Text)) <> 0 Then
        
        'verifica se o tamanho � maior do que zero
        If CInt(Tamanho.Text) < 1 Then Error 44900
        
        If Len(Trim(Tamanho.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
           objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
               
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44901

    Saida_Celula_Tamanho = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tamanho:

    Saida_Celula_Tamanho = Err
    
    Select Case Err
    
        Case 44900
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
             Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_INVALIDO", Err)
    
        Case 44901
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174384)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long
'faz a critica da celula tipo do grid que est� deixando de ser a corrente
'se for preenchido, o numero de linhas existentes no grid aumenta uma unidade

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo
    
    If Len(Trim(Tipo.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44902

    Saida_Celula_Tipo = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = Err
    
    Select Case Err
    
        Case 44902
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174385)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Preenchimento(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Preenchimento

    Set objGridInt.objControle = Preenchimento
                
    If Len(Trim(Preenchimento.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44903

    Saida_Celula_Preenchimento = SUCESSO
    
Exit Function
    
Erro_Saida_Celula_Preenchimento:

    Saida_Celula_Preenchimento = Err
    
    Select Case Err
    
        Case 44903
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174386)
        
    End Select

    Exit Function

End Function

Function Inicializa_Grid_Segmento() As Long
   
    'tela em quest�o
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGrid1.colColuna.Add ("Segmento")
    objGrid1.colColuna.Add ("Tipo")
    objGrid1.colColuna.Add ("Tamanho")
    objGrid1.colColuna.Add ("Delimitador")
    objGrid1.colColuna.Add ("Preenchimento")
    
   'campos de edi��o do grid
    objGrid1.colCampo.Add (Tipo.Name)
    objGrid1.colCampo.Add (Tamanho.Name)
    objGrid1.colCampo.Add (Delimitador.Name)
    objGrid1.colCampo.Add (Preenchimento.Name)
    
    objGrid1.objGrid = GridSegmentos
   
    'todas as linhas do grid
    objGrid1.objGrid.Rows = 10
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9
    
    objGrid1.objGrid.ColWidth(0) = 1000
    
    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGrid1)
    
    Inicializa_Grid_Segmento = SUCESSO
    
End Function

Function Grid_Segmentos(colSegmentos As Collection) As Long

Dim iIndice1 As Integer
Dim objsegmento As ClassSegmento

On Error GoTo Erro_Grid_Segmentos

    'percorre todas as linhas do grid
    For iIndice1 = 1 To objGrid1.iLinhasExistentes

        Set objsegmento = New ClassSegmento
                     
        'verifica se foi preenchido o campo formato
        If Len(Trim(Formato.Text)) = 0 Then Error 44904
        
        'inclui o Formato(codigo) em objSegmento
        objsegmento.sCodigo = sCodigo
              
        'inclui o nivel em objSegmento
        objsegmento.iNivel = iIndice1
        
        'verifica se foi preenchido o campo tipo
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_TIPO))) = 0 Then Error 44905
        
        'inclui o tipo em objSegmento
        objsegmento.iTipo = gobjColTipoSegmento.TipoSegmento(GridSegmentos.TextMatrix(iIndice1, COL_TIPO))
         
        'verifica se foi preenchido o campo tamanho
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_TAMANHO))) = 0 Then Error 44906
        
        'inclui o tamanho em objSegmento
        objsegmento.iTamanho = CInt(GridSegmentos.TextMatrix(iIndice1, COL_TAMANHO))
        
        'verifica se foi preenchido o campo delimitador
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_DELIMITADOR))) = 0 Then Error 44907
        
        'inclui o delimitador em objSegmento
        objsegmento.sDelimitador = GridSegmentos.TextMatrix(iIndice1, COL_DELIMITADOR)
        
        'verifica se foi preenchido o campo preenchimento
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_PREENCHIMENTO))) = 0 Then Error 44908
        
        'inclui o preenchimento em objSegmento
        objsegmento.iPreenchimento = gobjColPreenchimento.Preenchimento(GridSegmentos.TextMatrix(iIndice1, COL_PREENCHIMENTO))
        
        'Armazena o objeto objSegmento na cole��o colSegmento
        colSegmentos.Add objsegmento

    Next

    Grid_Segmentos = SUCESSO

    Exit Function

Erro_Grid_Segmentos:

    Grid_Segmentos = Err

    Select Case Err

        Case 44904
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_FORMATO_NAO_PREENCHIDO", Err)
            Formato.SetFocus
            
        Case 44905
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TIPO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_TIPO
            GridSegmentos.SetFocus
        
        Case 44906
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_TAMANHO
            GridSegmentos.SetFocus

        Case 44907
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_DELIMITADOR
            GridSegmentos.SetFocus

        Case 44908
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PREENCHIMENTO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_PREENCHIMENTO
            GridSegmentos.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174387)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SEGMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Segmentos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "SegmentosMAT"
    
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




Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

