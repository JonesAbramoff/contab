VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClasseUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'responsavel Jones
'revisada em 29/08/98
'pendencias:
    ' ver observacoes em MATGrava

Option Explicit

'Property Variables:
Dim m_objUserControl As Object
Dim m_Caption As String

Event Unload()

Dim iAlterado As Integer
Dim objGrid As AdmGrid

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Dim iGrid_Sigla_Col As Integer
Dim iGrid_Nome_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Conversao_Col As Integer
Dim iGrid_SiglaUMBase1_Col As Integer

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property


'Controles da Tela

Public Property Get Descricao() As Object
    Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get Codigo() As Object
    Set Codigo = objUserControl.Controls("Codigo")
End Property
Public Property Get LabelCodigo() As Object
    Set LabelCodigo = objUserControl.Controls("LabelCodigo")
End Property
Public Property Get SiglaUMBase() As Object
    Set SiglaUMBase = objUserControl.Controls("SiglaUMBase")
End Property
Public Property Get NomeUMBase() As Object
    Set NomeUMBase = objUserControl.Controls("NomeUMBase")
End Property
Public Property Get LabelConversao() As Object
    Set LabelConversao = objUserControl.Controls("LabelConversao")
End Property
Public Property Get Sigla() As Object
    Set Sigla = objUserControl.Controls("Sigla")
End Property
Public Property Get Nome() As Object
    Set Nome = objUserControl.Controls("Nome")
End Property
Public Property Get Quantidade() As Object
    Set Quantidade = objUserControl.Controls("Quantidade")
End Property
Public Property Get Conversao() As Object
    Set Conversao = objUserControl.Controls("Conversao")
End Property
Public Property Get SiglaUMBase1() As Object
    Set SiglaUMBase1 = objUserControl.Controls("SiglaUMBase1")
End Property
Public Property Get GridUM() As Object
    Set GridUM = objUserControl.Controls("GridUM")
End Property
Public Property Get BotaoGravar() As Object
    Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property
Public Property Get BotaoExcluir() As Object
    Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property
Public Property Get BotaoLimpar() As Object
    Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property
Public Property Get BotaoFechar() As Object
    Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

'Fim de Controles da Tela


Sub BotaoExcluir_Click()
'Chama ClasseUM_Excluir

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Código da ClasseUM foi preenchida
    If Len(Codigo.ClipText) = 0 Then Error 22411

    'Preenche objClasseUM
    objClasseUM.iClasse = CInt(Codigo.ClipText)
    objClasseUM.sDescricao = Descricao.Text
    objClasseUM.sSiglaUMBase = SiglaUMBase.Text

    'Envia aviso perguntando se realmente deseja excluir ClasseUM
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CLASSEUM", objClasseUM.iClasse)

    If vbMsgRes = vbYes Then

        'Exclui a Classe UM
        lErro = CF("ClasseUM_Exclui", objClasseUM)
        If lErro <> SUCESSO Then Error 22412

        'Limpa a Tela
        Call Limpar_Tela

    End If

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        'Erro já tratado
        Case 22412

        Case 22411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSEUM_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155530)

    End Select

    Exit Sub

End Sub

Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 57760
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57760 'Erro tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155531)
          
    End Select

    Exit Sub
    
End Sub

Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Form_Load()
'Coloca um default para a classe

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGrid = New AdmGrid

    Set objEventoNumero = New AdmEvento

    'Inicializa o Grid
    lErro = Inicializa_Grid_UM(objGrid)
    If lErro <> SUCESSO Then Error 22370
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22370

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155532)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_UM(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Sigla")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("X")
    objGridInt.colColuna.Add ("SiglaUMBase")

   'campos de edição do grid
    objGridInt.colCampo.Add (Sigla.Name)
    objGridInt.colCampo.Add (Nome.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Conversao.Name)
    objGridInt.colCampo.Add (SiglaUMBase1.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Sigla_Col = 1
    iGrid_Nome_Col = 2
    iGrid_Quantidade_Col = 3
    iGrid_Conversao_Col = 4
    iGrid_SiglaUMBase1_Col = 5

    objGridInt.objGrid = GridUM

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    GridUM.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    'para garantir que o "X" fique centralizado
    GridUM.ColAlignment(4) = flexAlignCenterCenter
    
    Inicializa_Grid_UM = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional objClasseUM As ClassClasseUM) As Long
'Se a classeUM vier preenchida colocá-la na tela

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma classe foi passada por parametro
    If Not (objClasseUM Is Nothing) Then

            'Lê a tabela ClasseUM a partir da Classe
            lErro = CF("ClasseUM_Le", objClasseUM)
            If lErro = SUCESSO Then

                lErro = Traz_ClasseUM_Tela(objClasseUM)
                If lErro <> SUCESSO Then Error 22386

            Else
            
                Codigo.Text = CStr(objClasseUM.iClasse)
                
            End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        'Erro  já tratado
        Case 22386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155533)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Traz_ClasseUM_Tela(objClasseUM As ClassClasseUM) As Long

Dim lErro As Long
Dim iIndice As Integer, iIndiceBase As Integer
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection

On Error GoTo Erro_Traz_ClasseUM_Tela

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGrid)

    'Exibe os dados de objClasseUM na tela
    Codigo.Text = CStr(objClasseUM.iClasse)
    Descricao.Text = objClasseUM.sDescricao
    SiglaUMBase.Text = objClasseUM.sSiglaUMBase
    
    'Lê a tabela UnidadesDeMedida a partir de ClasseUM
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then Error 22387

    'Exibe os dados da coleção na tela

    'obtem o indice da sigla base
    For iIndiceBase = 1 To colSiglas.Count

        If colSiglas.Item(iIndiceBase).sSigla = objClasseUM.sSiglaUMBase Then Exit For

    Next

    'Preenche NomeUMBase
    NomeUMBase.Text = colSiglas.Item(iIndiceBase).sNome

    'inclui a linha da sigla base como a 1a do grid
    GridUM.TextMatrix(1, iGrid_Sigla_Col) = colSiglas.Item(iIndiceBase).sSigla
    GridUM.TextMatrix(1, iGrid_Nome_Col) = colSiglas.Item(iIndiceBase).sNome
    GridUM.TextMatrix(1, iGrid_Quantidade_Col) = CStr(colSiglas.Item(iIndiceBase).dQuantidade)
    GridUM.TextMatrix(1, iGrid_SiglaUMBase1_Col) = colSiglas.Item(iIndiceBase).sSiglaUMBase
    GridUM.TextMatrix(1, iGrid_Conversao_Col) = "X"

    'remove a sigla base da colecao
    Call colSiglas.Remove(iIndiceBase)

    'inclui as outras siglas da classe no grid

    'Exibe os dados da coleção na tela
    For iIndice = 1 To colSiglas.Count

        'Insere no GridUM
        GridUM.TextMatrix(iIndice + 1, iGrid_Sigla_Col) = colSiglas.Item(iIndice).sSigla
        GridUM.TextMatrix(iIndice + 1, iGrid_Nome_Col) = colSiglas.Item(iIndice).sNome
        GridUM.TextMatrix(iIndice + 1, iGrid_Quantidade_Col) = CStr(colSiglas.Item(iIndice).dQuantidade)
        GridUM.TextMatrix(iIndice + 1, iGrid_SiglaUMBase1_Col) = colSiglas.Item(iIndice).sSiglaUMBase
        GridUM.TextMatrix(iIndice + 1, iGrid_Conversao_Col) = "X"

    Next

    objGrid.iLinhasExistentes = 1 + colSiglas.Count

    'Zerar iAlterado
    iAlterado = 0

    Traz_ClasseUM_Tela = SUCESSO

    Exit Function

Erro_Traz_ClasseUM_Tela:

    Traz_ClasseUM_Tela = Err

    Select Case Err

        Case 22387

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155534)

    End Select

    Exit Function

End Function

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22373

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case Err

        'Erros já tratados
        Case 22373

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155535)

    End Select

End Sub

Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22378

    'Limpa a tela
    Call Limpar_Tela

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        'Erro já tratado
        Case 22378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155536)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Chama ClasseUM_Grava

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 22374

    'Verifica se a Descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then Error 22375

    'Verifica se a Sigla está preenchida
    If Len(Trim(SiglaUMBase.Text)) = 0 Then Error 22376

    'Verifica se o Nome está preenchido
    If Len(Trim(NomeUMBase.Text)) = 0 Then Error 22377

    'Chama Move_Tela_Memoria para passar os dados da tela para  os objetos
    lErro = Move_Tela_Memoria(objClasseUM, colSiglas)
    If lErro <> SUCESSO Then Error 22393

    lErro = Trata_Alteracao(objClasseUM, objClasseUM.iClasse)
    If lErro <> SUCESSO Then Error 32313

    'Chama a função de gravação
    lErro = CF("ClasseUM_Grava", objClasseUM, colSiglas)
    If lErro <> SUCESSO Then Error 22394

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        'Erro já tratado
        Case 22393, 22394, 32313

        Case 22374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSEUM_NAO_PREENCHIDO", Err)

        Case 22375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_CLASSEUM_NAO_PREENCHIDA", Err)

        Case 22376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAUMBASE_NAO_PREENCHIDA", Err)

        Case 22377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEUMBASE_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155537)

    End Select

End Function

'Chama a função que limpa toda a tela
Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22526

    Call Limpar_Tela

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 22526

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155538)

    End Select

End Sub

'Usada para limpar a Tela toda
Sub Limpar_Tela()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa GridUM
    Call Grid_Limpa(objGrid)

    'Linhas visíveis do grid
    objGrid.iLinhasExistentes = 0
    
    iAlterado = 0

    Exit Sub

End Sub



Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
    Set objGrid = Nothing
    Set objEventoNumero = Nothing
    
End Sub

Sub GridUM_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Sub GridUM_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Sub GridUM_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridUM_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridUM_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridUM_KeyDown

    'A primeira linha não pode ser excluída
    If GridUM.Row = 1 And vbKeyDelete = KeyCode Then Error 22462

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    Exit Sub

Erro_GridUM_KeyDown:

    Select Case Err

        Case 22462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRIMEIRA_LINHA_NAO_PODE_SER_EXCLUIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155539)

    End Select

    Exit Sub

End Sub

Sub GridUM_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridUM_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Sub GridUM_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Sub GridUM_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridUM.Col

            Case iGrid_Sigla_Col

                lErro = Saida_Celula_Sigla(objGridInt)
                If lErro <> SUCESSO Then Error 22379

            Case iGrid_Nome_Col

                lErro = Saida_Celula_Nome(objGridInt)
                If lErro <> SUCESSO Then Error 22380

            Case iGrid_Quantidade_Col

                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then Error 22381

            Case iGrid_Conversao_Col

                lErro = Saida_Celula_Conversao(objGridInt)
                If lErro <> SUCESSO Then Error 22464

            Case iGrid_SiglaUMBase1_Col

                lErro = Saida_Celula_SiglaUMBase1(objGridInt)
                If lErro <> SUCESSO Then Error 22382

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 22383

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 22383
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22379, 22380, 22381, 22382, 22464

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155540)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Sigla(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Sigla

    Set objGridInt.objControle = Sigla

    If Len(Trim(Sigla.Text)) <> 0 And GridUM.Row - GridUM.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        iIndice = objGridInt.iLinhasExistentes

        'Seta o valor das duas últimas colunas automaticamente, quando uma nova linha é criada
        GridUM.TextMatrix(iIndice, iGrid_Conversao_Col) = "X"
        GridUM.TextMatrix(iIndice, iGrid_SiglaUMBase1_Col) = SiglaUMBase.Text

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22384

    Saida_Celula_Sigla = SUCESSO

    Exit Function

Erro_Saida_Celula_Sigla:

    Saida_Celula_Sigla = Err

    Select Case Err

        Case 22384
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155541)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Nome(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Nome do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Nome

    Set objGridInt.objControle = Nome

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22385

    Saida_Celula_Nome = SUCESSO

    Exit Function

Erro_Saida_Celula_Nome:

    Saida_Celula_Nome = Err

    Select Case Err

        Case 22385
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155542)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then Error 22461

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 19328

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = Err

    Select Case Err

        Case 22461, 22385, 19328
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155543)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conversao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Conversao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Conversao

    Set objGridInt.objControle = Conversao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22463

    Saida_Celula_Conversao = SUCESSO

    Exit Function

Erro_Saida_Celula_Conversao:

    Saida_Celula_Conversao = Err

    Select Case Err

        Case 22463
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155544)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SiglaUMBase1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula SiglaUMBase1 do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_SiglaUMBase1

    Set objGridInt.objControle = SiglaUMBase1

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22385

    Saida_Celula_SiglaUMBase1 = SUCESSO

    Exit Function

Erro_Saida_Celula_SiglaUMBase1:

    Saida_Celula_SiglaUMBase1 = Err

    Select Case Err

        Case 22385
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155545)

    End Select

    Exit Function

End Function

Sub LabelCodigo_Click()

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSelecao As New Collection
Dim colSiglas As New Collection

On Error GoTo Erro_LabelCodigo_Click

    If Len(Trim(Codigo.ClipText)) > 0 Then
        objClasseUM.iClasse = CInt(Codigo.ClipText)
    Else
        objClasseUM.iClasse = 0
    End If

    'Chama tela
    Call Chama_Tela("ClasseUMLista", colSelecao, objClasseUM, objEventoNumero)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case Err

        'Erro  já  tratado
        Case 22402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155546)

    End Select

    Exit Sub

End Sub

Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objClasseUM As ClassClasseUM

    Set objClasseUM = obj1

    'Chama Traz_ClasseUM_Tela
    lErro = Traz_ClasseUM_Tela(objClasseUM)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

End Sub

Sub NomeUMBase_Change()

    iAlterado = REGISTRO_ALTERADO

    GridUM.TextMatrix(1, iGrid_Nome_Col) = NomeUMBase.Text

End Sub

Sub SiglaUMBase_Change()
Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    GridUM.TextMatrix(1, iGrid_Sigla_Col) = SiglaUMBase.Text
    GridUM.TextMatrix(1, iGrid_SiglaUMBase1_Col) = SiglaUMBase.Text

    For iLinha = 2 To objGrid.iLinhasExistentes
        GridUM.TextMatrix(iLinha, iGrid_SiglaUMBase1_Col) = SiglaUMBase.Text
    Next

End Sub

Sub SiglaUMBase_Validate(Cancel As Boolean)

    'Verifica se a Sigla está preenchida
    If Len(Trim(SiglaUMBase.Text)) > 0 Then

       If objGrid.iLinhasExistentes = 0 Then Call Criar_Linha1
            
    End If

End Sub

Private Sub Criar_Linha1()

        objGrid.iLinhasExistentes = 1

        'Criar a primeira linha do Grid com conversão da UM para ela mesma
        GridUM.TextMatrix(1, iGrid_Sigla_Col) = SiglaUMBase.Text
        GridUM.TextMatrix(1, iGrid_Nome_Col) = NomeUMBase.Text
        GridUM.TextMatrix(1, iGrid_Quantidade_Col) = 1
        GridUM.TextMatrix(1, iGrid_Conversao_Col) = "X"
        GridUM.TextMatrix(1, iGrid_SiglaUMBase1_Col) = SiglaUMBase.Text

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As New ClassUnidadeDeMedida

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ClasseUM"

    If Len(Trim(Codigo.ClipText)) > 0 Then objClasseUM.iClasse = CInt(Codigo.ClipText)
    If Len(Trim(Descricao.Text)) > 0 Then objClasseUM.sDescricao = Descricao.Text

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Classe", objClasseUM.iClasse, 0, "Classe"
    colCampoValor.Add "Descricao", objClasseUM.sDescricao, STRING_CLASSEUM_DESCRICAO, "Descricao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        'Erros já tratados
        Case 22388

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155547)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objClasseUM As ClassClasseUM, colSiglas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objUnidadeDeMedida As New ClassUnidadeDeMedida

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objClasseUM
    If Len(Trim(Codigo.ClipText)) > 0 Then objClasseUM.iClasse = CInt(Codigo.ClipText)
    If Len(Trim(Descricao.Text)) > 0 Then objClasseUM.sDescricao = Descricao.Text
    If Len(Trim(SiglaUMBase.Text)) > 0 Then objClasseUM.sSiglaUMBase = SiglaUMBase.Text

    'Ir preenchendo uma coleção com todas as linhas "existentes" do grid
    'A Sigla, a Quantidade e a SiglaUMBase tem que estar preenchidos, senão erro
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Sigla, a Quantidade e a SiglaUMBase estão preenchidos
        If Len(Trim(GridUM.TextMatrix(iIndice, iGrid_Sigla_Col))) = 0 Then Error 22389
        If Len(Trim(GridUM.TextMatrix(iIndice, iGrid_Nome_Col))) = 0 Then Error 22390
        If Len(Trim(GridUM.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then Error 29828
        If Len(Trim(GridUM.TextMatrix(iIndice, iGrid_SiglaUMBase1_Col))) = 0 Then Error 22391

        Set objUnidadeDeMedida = New ClassUnidadeDeMedida

        objUnidadeDeMedida.iClasse = CInt(Codigo.Text)
        objUnidadeDeMedida.sSigla = GridUM.TextMatrix(iIndice, iGrid_Sigla_Col)
        objUnidadeDeMedida.sNome = GridUM.TextMatrix(iIndice, iGrid_Nome_Col)
        objUnidadeDeMedida.dQuantidade = CDbl(GridUM.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objUnidadeDeMedida.sSiglaUMBase = GridUM.TextMatrix(iIndice, iGrid_SiglaUMBase1_Col)

        'Verifica se UM já existe
        For iIndice1 = 1 To colSiglas.Count

            If objUnidadeDeMedida.sSigla = colSiglas.Item(iIndice1).sSigla Then Error 22397

        Next

        colSiglas.Add objUnidadeDeMedida

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 22389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_CLASSEUM_NAO_PREENCHIDA", Err)

        Case 22397
             lErro = Rotina_Erro(vbOKOnly, "ERRO_UM_REPETIDA", Err)

        Case 22390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEUMBASE_NAO_PREENCHIDO", Err)

        Case 22391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAUMBASE_NAO_PREENCHIDA", Err)

        Case 29828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_CLASSEUM_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155548)

    End Select

    Exit Function

End Function

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM

On Error GoTo Erro_Tela_Preenche

    objClasseUM.iClasse = colCampoValor.Item("Classe").vValor

    If objClasseUM.iClasse > 0 Then

        'Lê a tabela ClasseUM a partir da Classe
        lErro = CF("ClasseUM_Le", objClasseUM)
        If lErro <> SUCESSO Then Error 22415

       'Traz dados da ClasseUM para a Tela
        lErro = Traz_ClasseUM_Tela(objClasseUM)
        If lErro <> SUCESSO Then Error 22392

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        'Erro já tratado
        Case 22392

        Case 22415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_INEXISTENTE", Err, objClasseUM.iClasse)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155549)

    End Select

    Exit Sub

End Sub

Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Nome_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Sub Nome_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Sub Nome_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Nome
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Sigla_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Sub Sigla_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Sub Sigla_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Sigla
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub SiglaUMBase1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub SiglaUMBase1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Sub SiglaUMBase1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Sub SiglaUMBase1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = SiglaUMBase1
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Sub Conversao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Conversao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Sub Conversao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Sub Conversao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Conversao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Só executa para entrada de célula
    If iLocalChamada = ROTINA_GRID_ENTRADA_CELULA Or iLocalChamada = ROTINA_GRID_CLICK Or iLocalChamada = ROTINA_GRID_TRATA_TECLA_CAMPO2 Then

        If GridUM.Row = 1 Or SiglaUMBase.Text = "" Then
            objControl.Enabled = False
        Else
            If objControl.Name <> "Conversao" And objControl.Name <> "SiglaUMBase1" Then
                If objControl.Name = "Sigla" Then
                    objControl.Enabled = True
                Else
                    If Len(Trim(GridUM.TextMatrix(GridUM.Row, iGrid_Sigla_Col))) <> 0 Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                End If
            End If
        End If

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 155550)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CLASSE_UNIDADE_MEDIDA
    Set Form_Load_Ocx = Me
    Caption = "Classe de Unidades de Medida"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ClasseUM"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get objUserControl() As Object
     Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    End If

End Sub

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property
