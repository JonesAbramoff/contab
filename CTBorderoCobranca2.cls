VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBorderoCobranca2"
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
Dim gobjBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao

Public objGrid As AdmGrid
Public iGrid_Cliente_Col As Integer
Dim iGrid_NomeRedCliente_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_NumTitulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Incluir_Col As Integer
Dim iGrid_Vencimento_Col As Integer

'###############################
'Inserido por Wagner
Const ORDENACAO_NUMTITULO = 0
Const ORDENACAO_CLIENTE = 1
Const ORDENACAO_VENCIMENTO = 2
'###############################

'--- inicio dos properties get dos controles da tela
Public Property Get Ordenacao() As Object
     Set Ordenacao = objUserControl.Controls("Ordenacao")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoVoltar() As Object
     Set BotaoVoltar = objUserControl.Controls("BotaoVoltar")
End Property

Public Property Get BotaoSeguir() As Object
     Set BotaoSeguir = objUserControl.Controls("BotaoSeguir")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get BotaoDocOriginal() As Object
     Set BotaoDocOriginal = objUserControl.Controls("BotaoDocOriginal")
End Property

Public Property Get BotaoDesmarcar() As Object
     Set BotaoDesmarcar = objUserControl.Controls("BotaoDesmarcar")
End Property

Public Property Get BotaoMarcar() As Object
     Set BotaoMarcar = objUserControl.Controls("BotaoMarcar")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get QtdParcelas() As Object
     Set QtdParcelas = objUserControl.Controls("QtdParcelas")
End Property

Public Property Get TotalParcelas() As Object
     Set TotalParcelas = objUserControl.Controls("TotalParcelas")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get QtdParcelasSelecionadas() As Object
     Set QtdParcelasSelecionadas = objUserControl.Controls("QtdParcelasSelecionadas")
End Property

Public Property Get TotalParcelasSelecionadas() As Object
     Set TotalParcelasSelecionadas = objUserControl.Controls("TotalParcelasSelecionadas")
End Property

Public Property Get CheckIncluir() As Object
     Set CheckIncluir = objUserControl.Controls("CheckIncluir")
End Property

Public Property Get DataVencto() As Object
     Set DataVencto = objUserControl.Controls("DataVencto")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get GridBorderoCobranca2() As Object
     Set GridBorderoCobranca2 = objUserControl.Controls("GridBorderoCobranca2")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Nome() As Object
     Set Nome = objUserControl.Controls("Nome")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoDesmarcar_Click()
'Desmarca todas as parcelas marcadas no Grid

Dim iLinha As Integer
Dim objInfoParcRec As New ClassInfoParcRec
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela a parcela em questão
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col) = INCLUIR_NAO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcRec = gobjBorderoCobrancaEmissao.colInfoParcRec.Item(iLinha)
        
        'Desmarca no Obj a parcela em questão
        objInfoParcRec.iMarcada = 0
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Limpa na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdParcelasSelecionadas.Caption = CStr(0)
    TotalParcelasSelecionadas.Caption = CStr(Format(0, "Standard"))

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objInfoParcRec As New ClassInfoParcRec
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridBorderoCobranca2.Row = 0 Then Error 60486
        
    'Se foi selecionada uma linha que está preenchida
    If GridBorderoCobranca2.Row <= objGrid.iLinhasExistentes Then
        
        Set objInfoParcRec = gobjBorderoCobrancaEmissao.colInfoParcRec.Item(GridBorderoCobranca2.Row)
               
        objParcelaReceber.lNumIntDoc = objInfoParcRec.lNumIntParc
        
        'Le o NumInterno do Titulo para passar no objTituloRec
        lErro = CF("ParcelaReceber_Le", objParcelaReceber)
        If lErro <> SUCESSO And lErro <> 19147 Then Error 60487
        
        'Se não encontrou a Parcela --> ERRO
        If lErro = 19147 Then Error 60488
        
        objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
        
        'Le os Dados do Titulo
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then Error 60489
        
        'Se não encontrou
        If lErro = 26061 Then Error 60490
        
        Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case Err
    
        Case 60486
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
         
        Case 60487, 60489  'Tratado na rotina chamada
        
        Case 60488
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)
        
        Case 60490
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155502)

    End Select

    Exit Sub
  
End Sub

Public Sub BotaoFechar_Click()
    
    'Fechar a tela
    Unload Me

End Sub

Public Sub BotaoMarcar_Click()
'Marca todas as parcelas no Grid

Dim iLinha As Integer
Dim dTotalParcelasSelecionadas As Double
Dim iNumParcelasSelecionadas As Integer
Dim objInfoParcRec As ClassInfoParcRec

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela a parcela em questão
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col) = INCLUIR_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcRec = gobjBorderoCobrancaEmissao.colInfoParcRec.Item(iLinha)
        
        'Marca no Obj a parcela em questão
        objInfoParcRec.iMarcada = 1
        
        'Faz o somatório da Qtd e Total das Parcelas Selecionadas
        dTotalParcelasSelecionadas = dTotalParcelasSelecionadas + CDbl(GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Valor_Col))
        iNumParcelasSelecionadas = iNumParcelasSelecionadas + 1
    
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Atualiza na tela os campos Qtd de Parcelas selecionadas e Valor total das Parcelas selecionadas
    QtdParcelasSelecionadas.Caption = CStr(iNumParcelasSelecionadas)
    TotalParcelasSelecionadas.Caption = CStr(Format(dTotalParcelasSelecionadas, "Standard"))

End Sub

Public Sub BotaoSeguir_Click()
    
    gobjBorderoCobrancaEmissao.dValorParcelasSelecionadas = CDbl(TotalParcelasSelecionadas.Caption)
    gobjBorderoCobrancaEmissao.iQtdeParcelasSelecionadas = CInt(QtdParcelasSelecionadas.Caption)
    
    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoCobranca3", gobjBorderoCobrancaEmissao)
                        
    'Fecha a tela
    Unload Me
        
End Sub

Public Sub BotaoVoltar_Click()

Dim lErro As Long
Dim dValor As Double

On Error GoTo Erro_BotaoVoltar_Click

    If Len(Trim(TotalParcelasSelecionadas.Caption)) <> 0 Then
        dValor = CDbl(TotalParcelasSelecionadas.Caption)
    End If
    
    If gobjBorderoCobrancaEmissao.dValorParcelasSelecionadas <> dValor Then
        gobjBorderoCobrancaEmissao.dValorParcelasSelecionadas = dValor
    End If
    
    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoCobranca", gobjBorderoCobrancaEmissao)

    'Limpa o grid da tela em questão
    Call Grid_Limpa(objGrid)

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoVoltar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155503)

    End Select

    Exit Sub

End Sub

Public Sub CheckIncluir_Click()

Dim iLinha As Integer
Dim dTotalParcelasSelecionadas As Double
Dim iNumParcelasSelecionadas As Integer
Dim objInfoParcRec As ClassInfoParcRec

    'Passa o número da linha que foi clickada para iLinha
    iLinha = GridBorderoCobranca2.Row
    
    'Passa a linha do Grid para o Obj
    Set objInfoParcRec = gobjBorderoCobrancaEmissao.colInfoParcRec.Item(iLinha)
    
    'Passa para o Obj se a parcela em questão foi marcada ou desmarcada
    objInfoParcRec.iMarcada = CInt(GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col))
                    
    'Se a parcela foi marcada
    If GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col) = INCLUIR_CHECADO Then

        'Acrescenta a nova parcela no somatório de Qtd de Parcelas selecionadas e Valor total de Parcelas selecionadas
        QtdParcelasSelecionadas.Caption = CStr(CInt(QtdParcelasSelecionadas.Caption) + 1)
        TotalParcelasSelecionadas.Caption = CStr(Format(CDbl(TotalParcelasSelecionadas.Caption) + CDbl(GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Valor_Col)), "Standard"))
        
    Else
    
        'Subtrai a parcela do somatório de Qtd de Parcelas selecionadas e Valor total de Parcelas selecionadas
        QtdParcelasSelecionadas.Caption = CStr(CInt(QtdParcelasSelecionadas.Caption) - 1)
        TotalParcelasSelecionadas.Caption = CStr(Format(CDbl(TotalParcelasSelecionadas.Caption) - CDbl(GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Valor_Col)), "Standard"))
        
    End If

End Sub

Public Sub CheckIncluir_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub CheckIncluir_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
    
End Sub

Public Sub CheckIncluir_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = NumTitulo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
      
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155504)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente.

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa variáveis para saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Finaliza variáveis para saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 15834

    End If
       
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
    
        Case 15834
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155505)
        
    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid = Nothing
    
    Set gobjBorderoCobrancaEmissao = Nothing
    
End Sub

Public Sub GridBorderoCobranca2_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If
    
    If GridBorderoCobranca2.Row <> 0 Then
        Call CF2(Me, "BC_Traz_RazaoSocial_Tela", StrParaLong(GridBorderoCobranca2.TextMatrix(GridBorderoCobranca2.Row, iGrid_Cliente_Col)))
    End If
    
End Sub

Public Sub GridBorderoCobranca2_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid)

End Sub

Public Sub GridBorderoCobranca2_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid, iAlterado)
    
End Sub

Public Sub GridBorderoCobranca2_LeaveCell()
    
    Call Saida_Celula(objGrid)
    
End Sub

Public Sub GridBorderoCobranca2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)
    
End Sub

Public Sub GridBorderoCobranca2_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub


Public Sub GridBorderoCobranca2_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub

Public Sub GridBorderoCobranca2_RowColChange()

    Call Grid_RowColChange(objGrid)
      
    If GridBorderoCobranca2.Row <> 0 Then
        Call CF2(Me, "BC_Traz_RazaoSocial_Tela", StrParaLong(GridBorderoCobranca2.TextMatrix(GridBorderoCobranca2.Row, iGrid_Cliente_Col)))
    End If
      
End Sub

Public Sub GridBorderoCobranca2_Scroll()

    Call Grid_Scroll(objGrid)
    
End Sub

Private Function Inicializa_Grid_BorderoCobranca2(objGridInt As AdmGrid, iRegistros As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_BorderoCobranca2
    
    'tela em questão
    Set objGrid.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Nº Título")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Incluir")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Nome.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (NumTitulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (DataVencto.Name)
    objGridInt.colCampo.Add (CheckIncluir.Name)
    
    iGrid_Cliente_Col = 1
    iGrid_NomeRedCliente_Col = 2
    iGrid_Filial_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_NumTitulo_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_Valor_Col = 7
    iGrid_Vencimento_Col = 8
    iGrid_Incluir_Col = 9
    
    objGridInt.objGrid = GridBorderoCobranca2
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    'todas as linhas do grid
    If objGridInt.iLinhasVisiveis >= iRegistros + 1 Then
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    Else
        objGridInt.objGrid.Rows = iRegistros + 1
    End If
    
    GridBorderoCobranca2.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_BorderoCobranca2 = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_BorderoCobranca2:

    Inicializa_Grid_BorderoCobranca2 = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155506)
        
    End Select

    Exit Function
        
End Function

Function Trata_Parametros(Optional objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
'Traz os dados das Parcelas a pagar para a Tela

Dim objInfoParcRec As ClassInfoParcRec
Dim iLinha As Integer, lErro As Long
Dim dTotal As Double
On Error GoTo Erro_Trata_Parametros

    Set gobjBorderoCobrancaEmissao = objBorderoCobrancaEmissao
    
    'Passa para a tela os dados das Parcelas Selecionadas
    QtdParcelasSelecionadas.Caption = CStr(gobjBorderoCobrancaEmissao.iQtdeParcelasSelecionadas)
    TotalParcelasSelecionadas.Caption = CStr(Format(gobjBorderoCobrancaEmissao.dValorParcelasSelecionadas, "Standard"))
            
    Set objGrid = New AdmGrid
        
    lErro = Inicializa_Grid_BorderoCobranca2(objGrid, gobjBorderoCobrancaEmissao.colInfoParcRec.Count)
    If lErro <> SUCESSO Then gError 14250
        
    '##########################################
    'Alterado por Wagner 23/03/2005
    
    Ordenacao.ListIndex = ORDENACAO_CLIENTE

    lErro = Traz_Parcelas_Tela()
    If lErro <> SUCESSO Then gError 132372

'    iLinha = 0
'
'    'Percorre todas as parcelas da Coleção passada por parâmetro
'    For Each objInfoParcRec In gobjBorderoCobrancaEmissao.colInfoParcRec
'
'        iLinha = iLinha + 1
'
'        'Passa para a tela os dados da parcela em questão
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Cliente_Col) = objInfoParcRec.lCliente
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_NomeRedCliente_Col) = objInfoParcRec.sNomeRedCliente
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcRec.iFilialCliente
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcRec.sSiglaDocumento
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_NumTitulo_Col) = objInfoParcRec.lNumTitulo
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcRec.iNumParcela
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objInfoParcRec.dValor, "Standard")
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objInfoParcRec.dtVencimento, "dd/mm/yyyy")
'        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col) = objInfoParcRec.iMarcada
'
'        'Soma ao total o valor da parcela em questão
'        dTotal = dTotal + objInfoParcRec.dValor
'
'    Next
'
'    'Passa para o Obj o número de parcelas passadas pela Coleção
'    objGrid.iLinhasExistentes = iLinha
'
'    'Passa para a tela o somatório da Qtd de Parcelas e do Número total de parcelas
'    QtdParcelas.Caption = CStr(objGrid.iLinhasExistentes)
'    TotalParcelas.Caption = CStr(Format(dTotal, "Standard"))
'
'    'Atualiza na tela os checkbox marcados
'    Call Grid_Refresh_Checkbox(objGrid)
'
'    iAlterado = 0
    '##########################################

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
     
    Select Case gErr
          
        Case 14250
        
        Case 132372 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155507)
     
    End Select
     
    Exit Function

End Function

Public Sub QtdParcelasSelecionadas_Change()

    BotaoSeguir.Enabled = (StrParaDbl(QtdParcelasSelecionadas.Caption) > 0)

End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
'Chama a tela de consulta de Títulos a Receber quando essa opção for selecionada no grid
    Call BotaoDocOriginal_Click
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_COBRANCA_P2
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Cobranca - Passo 2"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoCobranca2"
    
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then


    End If

End Sub

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

Function BC_Traz_RazaoSocial_Tela(ByVal objCT As Object, ByVal lCliente As Long) As Long
    BC_Traz_RazaoSocial_Tela = SUCESSO
End Function

'#############################################
'Inserido por Wagner
Public Sub Ordenacao_Click()

Dim lErro As Long

On Error GoTo Erro_Ordenacao_Click

    lErro = Traz_Parcelas_Tela()
    If lErro <> SUCESSO Then gError 132370

    Exit Sub
    
Erro_Ordenacao_Click:
    
    Select Case gErr
    
        Case 132370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155508)
     
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Parcelas_Tela() As Long

Dim colSaida As New Collection
Dim colCampos As New Collection
Dim objInfoParcRec As ClassInfoParcRec
Dim iLinha As Integer
Dim lErro As Long
Dim dTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_Traz_Parcelas_Tela

    GL_objMDIForm.MousePointer = vbHourglass

    If Ordenacao.ListIndex = ORDENACAO_NUMTITULO Then
        colCampos.Add "lNumTitulo"
        colCampos.Add "iNumParcela"
        colCampos.Add "lCliente"
        colCampos.Add "iFilialCliente"
    ElseIf Ordenacao.ListIndex = ORDENACAO_CLIENTE Then
        colCampos.Add "lCliente"
        colCampos.Add "iFilialCliente"
        colCampos.Add "lNumTitulo"
        colCampos.Add "iNumParcela"
    ElseIf Ordenacao.ListIndex = ORDENACAO_VENCIMENTO Then
        colCampos.Add "dtDataVencimentoReal"
        colCampos.Add "dtVencimento"
    Else
        colCampos.Add "dValor"
    End If
    
    lErro = Ordena_Colecao(gobjBorderoCobrancaEmissao.colInfoParcRec, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 132371
    
    iLinha = 0

    'Percorre todas as parcelas da Coleção passada por parâmetro
    For Each objInfoParcRec In colSaida

        iLinha = iLinha + 1

        'Passa para a tela os dados da parcela em questão
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Cliente_Col) = objInfoParcRec.lCliente
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_NomeRedCliente_Col) = objInfoParcRec.sNomeRedCliente
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcRec.iFilialCliente
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcRec.sSiglaDocumento
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_NumTitulo_Col) = objInfoParcRec.lNumTitulo
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcRec.iNumParcela
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objInfoParcRec.dValor, "Standard")
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objInfoParcRec.dtVencimento, "dd/mm/yyyy")
        GridBorderoCobranca2.TextMatrix(iLinha, iGrid_Incluir_Col) = objInfoParcRec.iMarcada
        
        'Soma ao total o valor da parcela em questão
        dTotal = dTotal + objInfoParcRec.dValor
        
    Next
        
    '######################
    'Acerto
    For iIndice = gobjBorderoCobrancaEmissao.colInfoParcRec.Count To 1 Step -1
        gobjBorderoCobrancaEmissao.colInfoParcRec.Remove (iIndice)
    Next
    For iIndice = 1 To colSaida.Count
        gobjBorderoCobrancaEmissao.colInfoParcRec.Add colSaida.Item(iIndice)
    Next
    '######################
        
    'Passa para o Obj o número de parcelas passadas pela Coleção
    objGrid.iLinhasExistentes = iLinha
    
    'Passa para a tela o somatório da Qtd de Parcelas e do Número total de parcelas
    QtdParcelas.Caption = CStr(objGrid.iLinhasExistentes)
    TotalParcelas.Caption = CStr(Format(dTotal, "Standard"))
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Traz_Parcelas_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Parcelas_Tela:

    GL_objMDIForm.MousePointer = vbDefault

    Traz_Parcelas_Tela = gErr
     
    Select Case gErr
    
        Case 132371
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155509)
     
    End Select
     
    Exit Function
    
End Function
'#############################################


Public Sub GridBorderoCobranca2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGrid.objTela = Me
        
        'Chama o Menu PopUp
        PopUpMenuGrid.PopupMenu PopUpMenuGrid.mnuGrid, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuGrid.objTela = Nothing
        
    End If

End Sub

