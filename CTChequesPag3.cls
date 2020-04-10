VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTChequesPag3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iGrid_Cheque_Col As Integer
Dim iGrid_Valor_Col As Integer
Public iGrid_Beneficiario_Col As Integer 'Alterado por Wagner
Dim iGrid_Atualizar_Col As Integer
Public objGridChequesPag3 As AdmGrid 'Alterado por Wagner
Public gobjChequesPag As ClassChequesPag 'Alterado por Wagner

'30/11/01 Marcelo
Dim iChequeImpresso As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoVoltar() As Object
     Set BotaoVoltar = objUserControl.Controls("BotaoVoltar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoSeguir() As Object
     Set BotaoSeguir = objUserControl.Controls("BotaoSeguir")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoImprimirAPartir() As Object
     Set BotaoImprimirAPartir = objUserControl.Controls("BotaoImprimirAPartir")
End Property

Public Property Get BotaoConfigurarImpressao() As Object
     Set BotaoConfigurarImpressao = objUserControl.Controls("BotaoConfigurarImpressao")
End Property

Public Property Get BotaoImprimirSelecao() As Object
     Set BotaoImprimirSelecao = objUserControl.Controls("BotaoImprimirSelecao")
End Property

Public Property Get BotaoImprimirTeste() As Object
     Set BotaoImprimirTeste = objUserControl.Controls("BotaoImprimirTeste")
End Property

Public Property Get BotaoImprimirTudo() As Object
     Set BotaoImprimirTudo = objUserControl.Controls("BotaoImprimirTudo")
End Property

Public Property Get BotaoNumAuto() As Object
     Set BotaoNumAuto = objUserControl.Controls("BotaoNumAuto")
End Property

Public Property Get Atualizar() As Object
     Set Atualizar = objUserControl.Controls("Atualizar")
End Property

Public Property Get Beneficiario() As Object
     Set Beneficiario = objUserControl.Controls("Beneficiario")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get Cheque() As Object
     Set Cheque = objUserControl.Controls("Cheque")
End Property

Public Property Get GridChequesPag3() As Object
     Set GridChequesPag3 = objUserControl.Controls("GridChequesPag3")
End Property

Public Property Get LabelConta() As Object
     Set LabelConta = objUserControl.Controls("LabelConta")
End Property

Public Property Get LabelQtdCheques() As Object
     Set LabelQtdCheques = objUserControl.Controls("LabelQtdCheques")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Atualizar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Public Sub Atualizar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Public Sub Atualizar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Atualizar
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Public Sub BotaoConfigurarImpressao_Click()

    Call Sist_ImpressoraDlg(1)

End Sub

Public Sub BotaoFechar_Click()
    
    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoImprimirAPartir_Click()

Dim iIndice As Integer, objInfoChequePag As ClassInfoChequePag
Dim iLinhaInicial As Integer, iLinha As Integer, lErro As Long

On Error GoTo Erro_BotaoImprimirAPartir_Click

    iLinhaInicial = 0
   
   'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridChequesPag3.iLinhasExistentes
        If GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
            iLinhaInicial = iLinha
            Exit For
        End If
    Next
    
    'uma linha tem que estar selecionada
    If iLinhaInicial = 0 Then Error 32124
    
    'desmarcar todos acima e marcar todos c/indice a partir da 1a linha marcada
    
    'percorre a colecao marcando todos os cheques selecionados e desmarcando os outros
    For iIndice = 1 To gobjChequesPag.ColInfoChequePag.Count
    
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iIndice)
        If iLinhaInicial <= iIndice Then
            objInfoChequePag.iImprimir = 1
        Else
            objInfoChequePag.iImprimir = 0
        End If
    Next
    
    'imprimir os cheques marcados na colecao
    Call ImprimeChequesSelecionados
    
    Exit Sub
    
Erro_BotaoImprimirAPartir_Click:

    Select Case Err

        Case 32124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_CHEQUES_MARCADOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155519)

    End Select

    Exit Sub
        
End Sub

Public Sub BotaoImprimirSelecao_Click()
    
Dim iIndice As Integer, objInfoChequePag As ClassInfoChequePag, iQtde As Integer

On Error GoTo Erro_BotaoImprimirSelecao_Click

    iQtde = 0
    
    'percorre a colecao marcando todos os cheques selecionados e desmarcando os outros
    For iIndice = 1 To gobjChequesPag.ColInfoChequePag.Count
    
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iIndice)
        If GridChequesPag3.TextMatrix(iIndice, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
            objInfoChequePag.iImprimir = 1
            iQtde = iQtde + 1
        Else
            objInfoChequePag.iImprimir = 0
        End If
    Next
    
    If iQtde > 0 Then
    
        'imprimir os cheques marcados na colecao
        Call ImprimeChequesSelecionados
        
    Else
    
        Error 56756
    
    End If

    Exit Sub
     
Erro_BotaoImprimirSelecao_Click:

    Select Case Err
          
        Case 56756
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECIONE_CHEQUES_NO_GRID", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155520)
     
    End Select
     
    Exit Sub

End Sub

Public Sub BotaoImprimirTeste_Click()

Dim lErro As Long, lNumImpressao As Long

On Error GoTo Erro_ImprimirTeste_Click

    lErro = CF("Cheques_PrepararTesteImpressao", lNumImpressao)
    If lErro <> SUCESSO Then Error 19464

    gobjChequesPag.lNumImpressao = lNumImpressao

    lErro = ImprimirCheques(lNumImpressao, gobjChequesPag.sLayoutCheque, gobjChequesPag.dtEmissao)
    If lErro <> SUCESSO Then Error 19465

    Exit Sub
    
Erro_ImprimirTeste_Click:

    Select Case Err

        Case 19464, 19465
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155521)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoImprimirTudo_Click()
    
Dim objInfoChequePag As ClassInfoChequePag

    'percorre a colecao marcando tudo
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag
    
        objInfoChequePag.iImprimir = 1
        
    Next
    
    'imprime os cheques marcados na colecao
    Call ImprimeChequesSelecionados
    
End Sub

Public Sub BotaoNumAuto_Click()
'Gera numeração automática de cheques a partir do cheque selecionado no grid

Dim lErro As Long
Dim iNumParcelasMarcadas As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim iLinhaMarcada As Integer
Dim lChequeMarcado As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_BotaoNumAuto_Click
    
   lChequeMarcado = 0
   iLinhaMarcada = 0
   iNumParcelasMarcadas = 0
   
   'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridChequesPag3.iLinhasExistentes
        
        'Se o Cheque está marcado
        If GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
    
            'Passa a linha do Grid para o Obj
            Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iLinha)
            
            'Lê os dados do Cheque marcado
            lChequeMarcado = objInfoChequePag.lNumRealCheque
            iLinhaMarcada = iLinha
            
            'Faz o somatório do número de Cheques marcadas
            iNumParcelasMarcadas = iNumParcelasMarcadas + 1
            
            'Desmarca o Cheque
            GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_NAO_CHECADO
            
        End If
        
    Next
    
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridChequesPag3)
    
    'Verifica se há um número diferente de 1 de Cheques marcados
    If iNumParcelasMarcadas <> 1 Then Error 15876
    
    'Verifica se o Cheque marcado é o último do Grid
    If iLinhaMarcada = objGridChequesPag3.iLinhasExistentes Then Exit Sub
    
    iLinha = iLinhaMarcada
    iIndice = 0
   
    'Percorre todos os Cheques a partir do que foi selecionado
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag

        iIndice = iIndice + 1
        iLinha = iLinha + 1

        If iLinha > objGridChequesPag3.iLinhasExistentes Then Exit For
        
        'Passa a linha do Grid para o Obj
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iLinha)
            
        'Altera sequencialmente a numeração do Cheque
        objInfoChequePag.lNumRealCheque = lChequeMarcado + CLng(iIndice)
        GridChequesPag3.TextMatrix(iLinha, iGrid_Cheque_Col) = objInfoChequePag.lNumRealCheque
        
    Next
                    
    Exit Sub
    
Erro_BotaoNumAuto_Click:

    Select Case Err
    
        Case 15876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_CHEQUES_MARCADOS", Err)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155522)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoSeguir_Click()

'30/11/01 Marcelo inicio
Dim vbMsgRes As VbMsgBoxResult
    
    If iChequeImpresso = 0 Then
    
        'perguntar se prossegue mesmo sem ter impresso
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NAO_IMP_CHEQUE")
        If vbMsgRes <> vbYes Then Exit Sub
        
    End If
    
'30/11/01 Marcelo fim

    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequesPag4", gobjChequesPag)
    
    'Fecha a tela
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
        
        
    '30/11/01 Marcelo
    iChequeImpresso = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155523)
    
    End Select
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do Grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Set gObjChequesPag = objChequesPag
    
    'Chama rotina de inicialização da saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt.objGrid Is GridChequesPag3 Then

            Select Case objGridInt.objGrid.Col
            
                'Se a célula for o campo Cheque
                Case iGrid_Cheque_Col
                    
                    Set objGridInt.objControle = Cheque
                    
                   'Chama função de tratamento de saída da célula Cheque
                    lErro = Saida_Celula_Cheque(objGridInt)
                    If lErro <> SUCESSO Then gError 15874
                    
                Case iGrid_Beneficiario_Col
                
                    Set objGridInt.objControle = Beneficiario
                    
                   'Chama função de tratamento de saída da célula Beneficiario
                    lErro = Saida_Celula_Beneficiario(objGridInt)
                    If lErro <> SUCESSO Then gError 15874
                               
            End Select
            
        End If
        
        '########################################
        'Inserido por Wagner
        lErro = CF2(Me, "CHQ_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 132560
        '########################################
        
        'Chama função de finalização da saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 15875

    End If

    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 15874, 15875, 132560 'Inserido por Wagner
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155524)
        
    End Select

    Exit Function

End Function

Public Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("ChequesPag2", gobjChequesPag)

    'Fecha a tela
    Unload Me
    
End Sub

Public Sub Cheque_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Public Sub Cheque_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Public Sub Cheque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Cheque
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridChequesPag3 = Nothing

    Set gobjChequesPag = Nothing
    
End Sub

Public Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)
      
End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Public Sub Beneficiario_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Public Sub Beneficiario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Public Sub Beneficiario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Beneficiario
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub GridChequesPag3_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridChequesPag3, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    End If
    
End Sub

Public Sub GridChequesPag3_GotFocus()
    
    Call Grid_Recebe_Foco(objGridChequesPag3)

End Sub

Public Sub GridChequesPag3_EnterCell()
    
    Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    
End Sub

Public Sub GridChequesPag3_LeaveCell()
    
    Call Saida_Celula(objGridChequesPag3)
    
End Sub

Public Sub GridChequesPag3_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridChequesPag3)
    
End Sub

Public Sub GridChequesPag3_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridChequesPag3, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    End If

End Sub

Public Sub GridChequesPag3_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridChequesPag3)

End Sub

Public Sub GridChequesPag3_RowColChange()

    Call Grid_RowColChange(objGridChequesPag3)
       
End Sub

Public Sub GridChequesPag3_Scroll()

    Call Grid_Scroll(objGridChequesPag3)
    
End Sub

Private Function Inicializa_Grid_ChequesPag3(objGridInt As AdmGrid, iNumCheques As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iProximo As Integer

On Error GoTo Erro_Inicializa_Grid_ChequesPag3
    
    'tela em questão
    Set objGridChequesPag3.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cheque")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Beneficiário")
    objGridInt.colColuna.Add ("Selecionar")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Cheque.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Beneficiario.Name)
    objGridInt.colCampo.Add (Atualizar.Name)
    
    iGrid_Cheque_Col = 1
    iGrid_Valor_Col = 2
    iGrid_Beneficiario_Col = 3
    iGrid_Atualizar_Col = 4
    
    '#######################################
    'Inserido por Wagner
    iProximo = 5 'ATULIZAR VALOR AO INSERIR CAMPOS ANTES DOS CAMPOS CUSTOMIZADOS
    
    'O iProximo volta com o valor do próximo da seqüência
    lErro = CF2(Me, "CHQ_Inicializa_Grid_ChequesPag3", objGridInt, iProximo)
    If lErro <> SUCESSO Then gError 132564
    '#######################################
        
    objGridInt.objGrid = GridChequesPag3
            
    'todas as linhas do grid
    If iNumCheques > objGridInt.iLinhasVisiveis Then
        objGridInt.objGrid.Rows = iNumCheques + 1
    Else
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    End If
    
    GridChequesPag3.ColWidth(0) = 400
        
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ChequesPag3 = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_ChequesPag3:

    Inicializa_Grid_ChequesPag3 = gErr
    
    Select Case gErr
    
        Case 132564 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155525)
        
    End Select

    Exit Function
        
End Function

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long
'Traz para a tela os dados dos Cheques marcados para emissão

Dim objInfoChequePag As ClassInfoChequePag
Dim iLinha As Integer, lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjChequesPag = objChequesPag
    
    'Passa a Conta Corrente para a tela
    LabelConta.Caption = CStr(gobjChequesPag.iCta)
    
    Set objGridChequesPag3 = New AdmGrid
        
    lErro = Inicializa_Grid_ChequesPag3(objGridChequesPag3, gobjChequesPag.ColInfoChequePag.Count)
    If lErro <> SUCESSO Then gError 19376
    
    iLinha = 0
   
    'Percorre todos os Cheques da Coleção passada por parâmetro
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag

        iLinha = iLinha + 1

        'Passa para a tela os dados do Cheque em questão
        GridChequesPag3.TextMatrix(iLinha, iGrid_Cheque_Col) = objInfoChequePag.lNumRealCheque
        GridChequesPag3.TextMatrix(iLinha, iGrid_Valor_Col) = CStr(Format(objInfoChequePag.dValor, "Standard"))
        GridChequesPag3.TextMatrix(iLinha, iGrid_Beneficiario_Col) = objInfoChequePag.sFavorecido
        GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_NAO_CHECADO
        
        '#####################################
        'Inserido por Wagner
        lErro = CF2(Me, "CHQ_Trata_Parametros", objInfoChequePag, iLinha)
        If lErro <> SUCESSO Then gError 132568
        '#####################################
        
    Next

    'Passa para o Obj o número de Cheques passados pela Coleção
    objGridChequesPag3.iLinhasExistentes = iLinha
    
    'Se o número de Cheques for maior que o número de linhas do Grid
    If iLinha + 1 > GridChequesPag3.Rows Then
        
        'Altera o número de linhas do Grid de acordo com o número de Cheques
        GridChequesPag3.Rows = iLinha + 1
    
    End If
    
    'Passa para a tela a Qtd de Cheques
    LabelQtdCheques.Caption = CStr(objGridChequesPag3.iLinhasExistentes)
        
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridChequesPag3)
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 19376, 132568 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155526)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Function Saida_Celula_Cheque(objGridInt As AdmGrid) As Long

Dim objInfoChequePag As ClassInfoChequePag
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Cheque

    If Not (gobjChequesPag Is Nothing) Then
    
        'Passa os dados da linha do Grid para o Obj
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(GridChequesPag3.Row)
            
        If Len(Trim(Cheque.Text)) <> 0 Then
        
            'Passa para o Obj o valor do Cheque que está na tela
            objInfoChequePag.lNumRealCheque = CLng(Trim(Cheque.Text))
            
        Else
        
            'Inicializa o Cheque no Obj
            objInfoChequePag.lNumRealCheque = 0
            
        End If
        
    End If
    
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15877
    
    Saida_Celula_Cheque = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Cheque:

    Saida_Celula_Cheque = Err
    
    Select Case Err

        Case 15877
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155527)
            
    End Select
    
    Exit Function
    
End Function

Public Function CHQPAG3_ImprimirCheques(ByVal objCT As Object, lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CHQPAG3_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 19467
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade, "TIGNORARMARGEM", "S")
    If lErro <> SUCESSO Then gError 7431

    CHQPAG3_ImprimirCheques = SUCESSO

    Exit Function

Erro_CHQPAG3_ImprimirCheques:

    CHQPAG3_ImprimirCheques = gErr

    Select Case gErr

        Case 7431, 19467

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155528)

    End Select

    Exit Function
    
End Function

Private Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim lErro As Long

On Error GoTo Erro_ImprimirCheques

    lErro = CF2(Me, "CHQPAG3_ImprimirCheques", lNumImpressao, sLayoutCheques, dtDataEmissao)
    If lErro <> SUCESSO Then gError 185055

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = gErr

    Select Case gErr
    
        Case 185055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185054)

    End Select

    Exit Function

End Function

Private Sub ImprimeChequesSelecionados()

Dim lErro As Long

On Error GoTo Erro_ImprimirTeste_Click

    lErro = CF("ChequesPag_PrepararImpressao", gobjChequesPag)
    If lErro <> SUCESSO Then Error 19307

    lErro = ImprimirCheques(gobjChequesPag.lNumImpressao, gobjChequesPag.sLayoutCheque, gobjChequesPag.dtEmissao)
    If lErro <> SUCESSO Then Error 19308
    
    '30/11/01 Marcelo
    iChequeImpresso = 1
    
    Exit Sub
    
Erro_ImprimirTeste_Click:

    Select Case Err

        Case 19307, 19308
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155529)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_IMPRESSAO_CHEQUES_P3
    Set Form_Load_Ocx = Me
    Caption = "Impressão de Cheques - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequesPag3"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

'***** fim do trecho a ser copiado ******
Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

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

'############################################
'Inserido por Wagner 18/04/2005
Public Function CHQ_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    CHQ_Saida_Celula = SUCESSO
End Function

Public Function CHQ_Inicializa_Grid_ChequesPag3(ByVal objCT As Object, objGridInt As AdmGrid, iProximo As Integer) As Long
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    CHQ_Inicializa_Grid_ChequesPag3 = SUCESSO
    
End Function

Public Function CHQ_Trata_Parametros(ByVal objCT As Object, ByVal objInfoChequePag As ClassInfoChequePag, ByVal iLinha As Integer) As Long
    CHQ_Trata_Parametros = SUCESSO
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

    
    End If
    
End Sub
'############################################

Private Function Saida_Celula_Beneficiario(objGridInt As AdmGrid) As Long

Dim objInfoChequePag As ClassInfoChequePag
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Beneficiario

    If Not (gobjChequesPag Is Nothing) Then
    
        'Passa os dados da linha do Grid para o Obj
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(GridChequesPag3.Row)
            
        'Passa para o Obj o beneficiario que está na tela
        objInfoChequePag.sFavorecido = Trim(Beneficiario.Text)
        
    End If
    
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15877
    
    Saida_Celula_Beneficiario = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Beneficiario:

    Saida_Celula_Beneficiario = Err
    
    Select Case Err

        Case 15877
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155527)
            
    End Select
    
    Exit Function
    
End Function



