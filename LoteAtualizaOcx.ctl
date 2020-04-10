VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl LoteAtualizaOcx 
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   ScaleHeight     =   5430
   ScaleWidth      =   7590
   Begin VB.TextBox Origem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "Origem"
      Top             =   3855
      Width           =   1485
   End
   Begin VB.TextBox ValorDoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5670
      TabIndex        =   15
      Text            =   "ValorDoc"
      Top             =   4185
      Width           =   1125
   End
   Begin VB.TextBox NumDoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4530
      TabIndex        =   14
      Text            =   "NumDoc"
      Top             =   4200
      Width           =   1065
   End
   Begin VB.CheckBox Atualiza 
      Height          =   210
      Left            =   975
      TabIndex        =   5
      Top             =   3885
      Width           =   825
   End
   Begin VB.TextBox Exercicio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3165
      TabIndex        =   7
      Text            =   "Exercicio"
      Top             =   3825
      Width           =   1080
   End
   Begin VB.TextBox Periodo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4245
      TabIndex        =   8
      Text            =   "Periodo"
      Top             =   3825
      Width           =   1065
   End
   Begin VB.TextBox Lote 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5385
      TabIndex        =   9
      Text            =   "Lote"
      Top             =   3825
      Width           =   825
   End
   Begin VB.CommandButton BotaoAtualizar 
      Caption         =   "Atualizar"
      Height          =   585
      Left            =   1965
      Picture         =   "LoteAtualizaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4725
      Width           =   1140
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todos"
      Height          =   570
      Left            =   1815
      Picture         =   "LoteAtualizaOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1425
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   570
      Left            =   3870
      Picture         =   "LoteAtualizaOcx.ctx":1174
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1425
   End
   Begin VB.CommandButton BotaoFechar 
      Caption         =   "Fechar"
      Height          =   585
      Left            =   3990
      Picture         =   "LoteAtualizaOcx.ctx":2356
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4725
      Width           =   1140
   End
   Begin VB.TextBox Status 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6165
      TabIndex        =   10
      Text            =   "Status"
      Top             =   3825
      Width           =   1065
   End
   Begin VB.CheckBox ExibirLotesAtualizando 
      Caption         =   "Exibir os lotes que estão sendo atualizados"
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   900
      Width           =   3495
   End
   Begin VB.ListBox auxExercicio 
      Height          =   255
      ItemData        =   "LoteAtualizaOcx.ctx":24D4
      Left            =   6000
      List            =   "LoteAtualizaOcx.ctx":24D6
      TabIndex        =   2
      Top             =   135
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.ListBox auxPeriodo 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   800
   End
   Begin MSFlexGridLib.MSFlexGrid GridLotesPendentes 
      Height          =   2805
      Left            =   120
      TabIndex        =   11
      Top             =   1125
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   4948
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
Attribute VB_Name = "LoteAtualizaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const COL_ATUALIZA = 1
Const COL_ORIGEM = 2
Const COL_EXERCICIO = 3
Const COL_PERIODO = 4
Const COL_LOTE = 5
Const COL_STATUS = 6
Const COL_NUMDOC = 7
Const COL_VALDOC = 8

Dim iAlterado As Integer
Dim objGrid1 As AdmGrid

Function Trata_Parametros() As Long

    iAlterado = 0
 
    Trata_Parametros = SUCESSO
 
    Exit Function
 
End Function

Private Sub BotaoDesmarcarTodos_Click()
    
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Error_BotaoDesmarcarTodos_Click

    'percorre todas as linhas do grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
    
        'marca cada checkbox Atualiza do grid
        GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "0"
    
    Next
    
    lErro = Grid_Refresh_Checkbox(objGrid1)
    If lErro <> SUCESSO Then Error 12233
    
    Exit Sub
    
Error_BotaoDesmarcarTodos_Click:

    Select Case Err
    
        Case 12233
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162441)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
    
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Error_BotaoMarcarTodos_Click

    'percorre todas as linhas do grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
    
        'marca cada checkbox Atualiza do grid
        GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "1"
    
    Next
    
    lErro = Grid_Refresh_Checkbox(objGrid1)
    If lErro <> SUCESSO Then Error 12232
    
    Exit Sub
    
Error_BotaoMarcarTodos_Click:

    Select Case Err
    
        Case 12232
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162442)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoAtualizar_Click()

Dim colLote As New Collection
Dim lErro As Long, sNomeArqParam As String
Dim iIDAtualizacao As Integer

On Error GoTo Error_BotaoAtualizar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'le os dados do grid que estao marcados para serem atualizados e coloca na colecao colLote
    lErro = GridLotesPendentes_Le(colLote)
    If lErro <> SUCESSO Then Error 12215
    
    'Atualiza o campo IdAtualizacao das tabelas Configuracao e LotePendente
    lErro = CF("LotePendente_Atualiza", colLote, iIDAtualizacao)
    If lErro <> SUCESSO Then Error 12230
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 27439
    
    lErro = CF("Rotina_Atualizacao", sNomeArqParam, iIDAtualizacao)
    If lErro <> SUCESSO Then Error 9407
    
    'limpa o grid
    Call Grid_Limpa(objGrid1)
       
    Set colLote = New Collection
    
    If ExibirLotesAtualizando.Value = 1 Then
    
        'le todos os lotes com status = desatualizado na tabela LotePendente e coloca na colecao colLote
        lErro = CF("LotePendente_Le_DesatualizadoII", giFilialEmpresa, colLote)
        If lErro <> SUCESSO Then Error 12231
    
    Else
    
        'leitura dos Lotes desatualizados e IdAtualizacao zerados
        lErro = CF("LotePendente_Le_Desatualizados", giFilialEmpresa, colLote)
        If lErro <> SUCESSO Then Error 12234
    
    End If
    
    'preenche o grid com os dados da colecao
    lErro = Grid_Preenche(colLote)
    If lErro <> SUCESSO Then Error 12235
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Error_BotaoAtualizar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case Err
    
        Case 9407, 12215, 12230, 12231, 12234, 12235, 27439
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162443)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ExibirLotesAtualizando_Click()

Dim colLote As New Collection
Dim lErro As Long

On Error GoTo Error_ExibirLotesAtualizando_Click

    'limpa o grid
    Call Grid_Limpa(objGrid1)
        
    If ExibirLotesAtualizando.Value = 1 Then
        
        'le todos os lotes com status = desatualizado na tabela LotePendente e coloca na colecao colLote
        lErro = CF("LotePendente_Le_DesatualizadoII", giFilialEmpresa, colLote)
        If lErro <> SUCESSO Then Error 12189
     
    Else
        
        'leitura dos Lotes desatualizados e IdAtualizacao zerados
        lErro = CF("LotePendente_Le_Desatualizados", giFilialEmpresa, colLote)
        If lErro <> SUCESSO Then Error 12237
    
    End If
    
    'preenche o grid com os dados da colecao colLote
    lErro = Grid_Preenche(colLote)
    If lErro <> SUCESSO Then Error 12229

    Exit Sub

Error_ExibirLotesAtualizando_Click:

    Select Case Err
    
        Case 12189, 12229, 12237
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162444)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim colLote As New Collection

On Error GoTo Erro_Form_Load

    Set objGrid1 = New AdmGrid
                   
    'Leitura dos Lotes desatualizados
    lErro = CF("LotePendente_Le_Desatualizados", giFilialEmpresa, colLote)
    If lErro <> SUCESSO Then Error 12213
    
    'se não encontrou lote desatualizado ==> pesquisa se há lotes em processo de atualização
    If colLote.Count = 0 Then
    
        'Tenta ler um lote pendente
        lErro = CF("LotePendente_Le1", giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 33611 Then Error 33609
        
        'Não encontrou nenhum lote pendente
        If lErro = 33611 Then Error 33610
        
        'Encontrou algum lote atualizando ==> avisa que só existem lotes sendo atualizados
        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_LOTE_ATUALIZANDO", giFilialEmpresa)
        
    End If
    
    'Inicializacao do grid
    Call Inicializa_Grid_LotesPendentes
     
    'Preenche o grid com os dados da colecao
    lErro = Grid_Preenche(colLote)
    If lErro <> SUCESSO Then Error 12214
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 12213, 12214, 33609
        
        Case 33610
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_LOTE_PENDENTE", Err, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162445)
            
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Sub Atualiza_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Private Sub Atualiza_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Atualiza_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Atualiza
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid1 = Nothing
    
End Sub

Private Sub Origem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Private Sub Origem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Origem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Origem
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Exercicio_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Exercicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Exercicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Exercicio
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Periodo_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Periodo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Periodo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Periodo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Lote_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Status_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Status_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Status
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridLotesPendentes_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridLotesPendentes_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridLotesPendentes_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridLotesPendentes_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridLotesPendentes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    
End Sub

Private Sub GridLotesPendentes_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridLotesPendentes_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridLotesPendentes_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridLotesPendentes_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

   lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 12195

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 12195
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162446)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_LotesPendentes() As Long
   
    'tela em questão
    Set objGrid1.objForm = Me
    
    objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Atualiza")
    objGrid1.colColuna.Add ("Origem")
    objGrid1.colColuna.Add ("Exercicio")
    objGrid1.colColuna.Add ("Periodo")
    objGrid1.colColuna.Add ("Lote")
    objGrid1.colColuna.Add ("Status")
    objGrid1.colColuna.Add ("Num. Docs.")
    objGrid1.colColuna.Add ("Valor Docs.")
    

   'campos de edição do grid
    objGrid1.colCampo.Add (Atualiza.Name)
    objGrid1.colCampo.Add (Origem.Name)
    objGrid1.colCampo.Add (Exercicio.Name)
    objGrid1.colCampo.Add (Periodo.Name)
    objGrid1.colCampo.Add (Lote.Name)
    objGrid1.colCampo.Add (Status.Name)
    objGrid1.colCampo.Add (NumDoc.Name)
    objGrid1.colCampo.Add (ValorDoc.Name)
    
    objGrid1.objGrid = GridLotesPendentes
   
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 9
    
    objGrid1.objGrid.ColWidth(0) = 500
    
    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL
        
    Call Grid_Inicializa(objGrid1)
     
    lErro_Chama_Tela = SUCESSO

    Inicializa_Grid_LotesPendentes = SUCESSO
    
End Function

Private Function Grid_Preenche(colLote As Collection) As Long
'preenche o GridLotesPendentes e as duas listbox invisiveis, com os dados da colecao colLote

Dim lErro As Long
Dim iLinha As Integer
Dim objLote As ClassLote
Dim objcolExercicio As New ClassColExercicio
Dim objcolPeriodo As New ClassColPeriodo

On Error GoTo Erro_Grid_Preenche
        
    If colLote.Count < 10 Then
        objGrid1.objGrid.Rows = 11
    Else
        objGrid1.objGrid.Rows = colLote.Count + 2
    End If
    
    objGrid1.iLinhasExistentes = colLote.Count
    
    'apenas inicializacao das listboxs para ficar relacionadas 1 a 1 com as linhas do grid
    auxExercicio.Clear
    auxPeriodo.Clear
    auxExercicio.AddItem "Exercicios"
    auxPeriodo.AddItem "Periodos"
    
    iLinha = 1
           
    'pega cada objeto da colecao para fazer os preenchimentos
    For Each objLote In colLote
    
        'preenche a listbox auxiliar iExercicio (invisivel na tela)
        auxExercicio.AddItem CStr(objLote.iExercicio)
        
        'preenche a listbox auxiliar iPeriodo (invisivel na tela)
        auxPeriodo.AddItem CStr(objLote.iPeriodo)
        
        'coloca a Origem no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_ORIGEM) = gobjColOrigem.Descricao(objLote.sOrigem)
        
        'coloca o Exercicio no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_EXERCICIO) = objcolExercicio.NomeExterno(objLote.iExercicio)

        'coloca o Periodo no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_PERIODO) = objcolPeriodo.NomeExterno(objLote.iExercicio, objLote.iPeriodo)

        'coloca o Lote no grid da tela
        GridLotesPendentes.TextMatrix(iLinha, COL_LOTE) = objLote.iLote
        
        'coloca o numero de documentos informado que compoe o lote na tela
        GridLotesPendentes.TextMatrix(iLinha, COL_NUMDOC) = objLote.iNumDocInf
        
        'coloca o valor dos documentos informado que compoe o lote na tela
        GridLotesPendentes.TextMatrix(iLinha, COL_VALDOC) = Format(objLote.dTotInf, "Standard")

        If objLote.iIDAtualizacao <> 0 Then

            'coloca o Status = Atualizando no grid da tela
            GridLotesPendentes.TextMatrix(iLinha, COL_STATUS) = LOTE_ATUALIZANDO_TEXTO

        End If
        
        iLinha = iLinha + 1
    
    Next

    Call Grid_Inicializa(objGrid1)

    Grid_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Preenche:

    Grid_Preenche = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162447)
            
    End Select
    
    Exit Function
    
End Function

Private Function GridLotesPendentes_Le(colLote As Collection) As Long
'Le o Grid e as listbox invisiveis aonde a checkbox Atualiza = 1, colocando os dados na colecao

Dim iIndice As Integer
Dim iTemLote As Integer
Dim objLote As ClassLote
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Error_GridLotesPendentes_Le

    iTemLote = 0

    'Percorre todas as linhas do grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
    
        'seleciona os registros marcados na checkbox atualiza
        If GridLotesPendentes.TextMatrix(iIndice, COL_ATUALIZA) = "1" Then
            
            iTemLote = 1
            
            Set objLote = New ClassLote
            
                objLote.iFilialEmpresa = giFilialEmpresa
            
                'insere a origem no objeto
                objLote.sOrigem = gobjColOrigem.Origem(GridLotesPendentes.TextMatrix(iIndice, COL_ORIGEM))
                
                'insere o exercicio no objeto
                objLote.iExercicio = CInt(auxExercicio.List(iIndice))
                
                'insere o periodo no objeto
                objLote.iPeriodo = CInt(auxPeriodo.List(iIndice))
                
                'insere o lote no objeto
                objLote.iLote = CInt(GridLotesPendentes.TextMatrix(iIndice, COL_LOTE))
                
                objLote.iNumDocInf = CInt(GridLotesPendentes.TextMatrix(iIndice, COL_NUMDOC))
                
                objLote.dTotInf = CDbl(GridLotesPendentes.TextMatrix(iIndice, COL_VALDOC))
                
                If GridLotesPendentes.TextMatrix(iIndice, COL_STATUS) = LOTE_ATUALIZANDO_TEXTO Then
                    objLote.iStatus = LOTE_ATUALIZANDO
                Else
                    objLote.iStatus = LOTE_DESATUALIZADO
                End If
                
                'adiciona o objeto a colecao
                colLote.Add objLote
        
        End If
    
    Next
    
    If iTemLote = 0 Then Error 12211
    
    GridLotesPendentes_Le = SUCESSO
    
Exit Function

Error_GridLotesPendentes_Le:

    GridLotesPendentes_Le = Err

    Select Case Err
                
        Case 12211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOTE", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162448)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_LOTE_ATUALIZA
    Set Form_Load_Ocx = Me
    Caption = "Atualização de Lotes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteAtualiza"
    
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


