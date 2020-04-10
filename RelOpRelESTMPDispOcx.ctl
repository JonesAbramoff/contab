VERSION 5.00
Begin VB.UserControl RelOpRelESTMPDispOcx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "RelOpRelESTMPDispOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub Form_Load()

Dim lErro As Long
Dim lNumIntRel As Long

On Error GoTo Erro_Form_Load
    
'    lErro = PreencherRelOp(gobjRelOpcoes)
'    If lErro <> SUCESSO Then gError 85182 '64963
'    Call CF("RelEstMPDisp_Prepara", lNumIntRel, giFilialEmpresa)
'    Call gobjRelatorio.Executar_Prossegue2(Me)

'    Set objEventoProdutoDe = New AdmEvento
'    Set objEventoProdutoAte = New AdmEvento
'
'    Set objEventoTipoInicial = New AdmEvento
'    Set objEventoTipoFinal = New AdmEvento
'
'    'Inicializa a mascara de produto
'    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
'    If lErro <> SUCESSO Then gError 85154
'
'    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
'    If lErro <> SUCESSO Then gError 85155
'
'    giProdInicial = 1
'
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 85154, 85155
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172569)

    End Select

    Exit Sub

End Sub

'Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
''l� os par�metros do arquivo C e exibe na tela
'
'Dim lErro As Long
'Dim sParam As String
'
'On Error GoTo Erro_PreencherParametrosNaTela
'
'    lErro = objRelOpcoes.Carregar
'    If lErro <> SUCESSO Then gError 85156
'
'    'pega Produto Inicial e exibe
'    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
'    If lErro <> SUCESSO Then gError 85157
'
'    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
'    If lErro <> SUCESSO Then gError 85158
'
'    'pega par�metro Produto Final e exibe
'    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
'    If lErro <> SUCESSO Then gError 85159
'
'    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
'    If lErro <> SUCESSO Then gError 85160
'
'    lErro = objRelOpcoes.ObterParametro("TTIPOPRODINI", sParam)
'    If lErro <> SUCESSO Then gError 85200
'
'    TipoInicial.Text = sParam
'
'    lErro = objRelOpcoes.ObterParametro("TTIPOPRODFIM", sParam)
'    If lErro <> SUCESSO Then gError 85201
'
'    TipoFinal.Text = sParam
'
'    PreencherParametrosNaTela = SUCESSO
'
'    Exit Function
'
'Erro_PreencherParametrosNaTela:
'
'    PreencherParametrosNaTela = gErr
'
'    Select Case gErr
'
'        Case 85156 To 85160, 85200, 85201
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172570)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim lNumIntRel As Long

On Error GoTo Erro_Trata_Parametros
    
    Call CF("RelEstMPDisp_Prepara", lNumIntRel, giFilialEmpresa)
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then gError 85177
    
    Call objRelatorio.Executar_Prossegue2(Me)


'    If Not (gobjRelatorio Is Nothing) Then gError 85161
'
'    Set gobjRelatorio = objRelatorio
'    Set gobjRelOpcoes = objRelOpcoes
'
'    'Preenche com as Opcoes
'    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
'    If lErro <> SUCESSO Then gError 85162
'
    Trata_Parametros = 1

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 85162, 85177
        
        Case 85161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172571)

    End Select

    Exit Function

End Function
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
''Formata os produtos retornando em sProd_I e sProd_F
''Verifica se os par�metros iniciais s�o maiores que os finais
'
'Dim iProdPreenchido_I As Integer
'Dim iProdPreenchido_F As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_Formata_E_Critica_Parametros
'
'    'formata o Produto Inicial
'    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
'    If lErro <> SUCESSO Then gError 85163
'
'    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""
'
'    'formata o Produto Final
'    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
'    If lErro <> SUCESSO Then gError 85164
'
'    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""
'
'    'se ambos os produtos est�o preenchidos, o produto inicial n�o pode ser maior que o final
'    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
'
'        If sProd_I > sProd_F Then gError 85165
'
'    End If
'
'    'tipo inicial n�o pode ser maior que o tipo final
'    If Trim(TipoInicial.Text) <> "" And Trim(TipoFinal.Text) <> "" Then
'
'         If Codigo_Extrai(TipoInicial.Text) > Codigo_Extrai(TipoFinal.Text) Then gError 85197
'
'    End If
'
'    Formata_E_Critica_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Formata_E_Critica_Parametros:
'
'    Formata_E_Critica_Parametros = gErr
'
'    Select Case gErr
'
'        Case 85163
'            ProdutoInicial.SetFocus
'
'        Case 85164
'            ProdutoFinal.SetFocus
'
'        Case 85165
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
'            ProdutoInicial.SetFocus
'
'        Case 85197
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", gErr)
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172572)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoLimpar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoLimpar_Click
'
'    lErro = Limpa_Relatorio(Me)
'    If lErro <> SUCESSO Then gError 85166
'
'    ComboOpcoes.Text = ""
'    DescProdInic.Caption = ""
'    DescProdFim.Caption = ""
'
'    TipoInicial.Text = ""
'    TipoFinal.Text = ""
'
'    giProdInicial = 1
'
'    ComboOpcoes.SetFocus
'
'    Exit Sub
'
'Erro_BotaoLimpar_Click:
'
'    Select Case gErr
'
'        Case 85166
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172573)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ComboOpcoes_Click()
'
'    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
'
'End Sub
'
'Private Sub ComboOpcoes_Validate(Cancel As Boolean)
'
'    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

'Private Sub LabelProdutoAte_Click()
'
'Dim lErro As Long
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim colSelecao As New Collection
'
'On Error GoTo Erro_LabelProdutoAte_Click
'
'    'Verifica se o produto foi preenchido
'    If Len(ProdutoFinal.ClipText) <> 0 Then
'
'        'Preenche o c�digo de objProduto
'        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 85167
'
'        objProduto.sCodigo = sProdutoFormatado
'
'    End If
'
'    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)
'
'    Exit Sub
'
'Erro_LabelProdutoAte_Click:
'
'    Select Case gErr
'
'        Case 85167
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172574)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objProduto As ClassProduto
'
'
'On Error GoTo Erro_objEventoProdutoDe_evSelecao
'
'    Set objProduto = obj1
'
'    'L� o Produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 85168
'
'    'Se n�o achou o Produto --> erro
'    If lErro = 28030 Then gError 85169
'
'    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
'    If lErro <> SUCESSO Then gError 85170
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoProdutoDe_evSelecao:
'
'    Select Case gErr
'
'        Case 85168, 85170
'
'        Case 85169
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172575)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objProduto As ClassProduto
'
'
'On Error GoTo Erro_objEventoProdutoAte_evSelecao
'
'    Set objProduto = obj1
'
'    'L� o Produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 85171
'
'    'Se n�o achou o Produto --> erro
'    If lErro = 28030 Then gError 85172
'
'    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
'    If lErro <> SUCESSO Then gError 85173
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoProdutoAte_evSelecao:
'
'    Select Case gErr
'
'        Case 85171, 85173
'
'        Case 85172
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172576)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LabelProdutoDe_Click()
'
'Dim lErro As Long
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim colSelecao As New Collection
'
'On Error GoTo Erro_LabelProdutoDe_Click
'
'    'Verifica se o produto foi preenchido
'    If Len(ProdutoInicial.ClipText) <> 0 Then
'
'        'Preenche o c�digo de objProduto
'        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 85174
'
'        objProduto.sCodigo = sProdutoFormatado
'
'    End If
'
'    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)
'
'    Exit Sub
'
'Erro_LabelProdutoDe_Click:
'
'    Select Case gErr
'
'        Case 85174
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172577)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
''preenche o arquivo C com os dados fornecidos pelo usu�rio
'
'Dim lErro As Long
'Dim sProd_I As String
'Dim sProd_F As String
'
'On Error GoTo Erro_PreencherRelOp
'
'    sProd_I = String(STRING_PRODUTO, 0)
'    sProd_F = String(STRING_PRODUTO, 0)
'
'    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
'    If lErro <> SUCESSO Then gError 85175
'
'    lErro = objRelOpcoes.Limpar
'    If lErro <> AD_BOOL_TRUE Then gError 85176
'
'    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
'    If lErro <> AD_BOOL_TRUE Then gError 85177
'
'    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
'    If lErro <> AD_BOOL_TRUE Then gError 85178
'
'    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODINI", TipoInicial.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 85198
'
'    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODFIM", TipoFinal.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 85199
'
'    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
'    If lErro <> SUCESSO Then gError 85179
'
'    PreencherRelOp = SUCESSO
'
'    Exit Function
'
'Erro_PreencherRelOp:
'
'    PreencherRelOp = gErr
'
'    Select Case gErr
'
'        Case 85175 To 85179, 85198, 85199
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172578)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoExcluir_Click()
'
'Dim vbMsgRes As VbMsgBoxResult
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoExcluir_Click
'
'    'verifica se nao existe elemento selecionado na ComboBox
'    If ComboOpcoes.ListIndex = -1 Then gError 85180 '64960
'
'    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")
'
'    If vbMsgRes = vbYes Then
'
'        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
'        If lErro <> SUCESSO Then gError 85181 '64961
'
'        'retira nome das op��es do ComboBox
'        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
'
'        'limpa as op��es da tela
'         lErro = Limpa_Relatorio(Me)
'        If lErro <> SUCESSO Then gError 85182 '64962
'
'        ComboOpcoes.Text = ""
'        DescProdInic.Caption = ""
'        DescProdFim.Caption = ""
'
'    End If
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    Select Case gErr
'
'        Case 85180
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
'            ComboOpcoes.SetFocus
'
'        Case 85181, 85182
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172579)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoExecutar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoExecutar_Click
'
'    lErro = PreencherRelOp(gobjRelOpcoes)
'    If lErro <> SUCESSO Then gError 85182 '64963
'
'    Call gobjRelatorio.Executar_Prossegue2(Me)
'
'    Exit Sub
'
'Erro_BotaoExecutar_Click:
'
'    Select Case gErr
'
'        Case 85182
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172580)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoGravar_Click()
''Grava a op��o de relat�rio com os par�metros da tela
'
'Dim lErro As Long
'Dim iResultado As Integer
'
'On Error GoTo Erro_BotaoGravar_Click
'
'    'nome da op��o de relat�rio n�o pode ser vazia
'    If ComboOpcoes.Text = "" Then gError 85183
'
'    lErro = PreencherRelOp(gobjRelOpcoes)
'    If lErro <> SUCESSO Then gError 85184
'
'    gobjRelOpcoes.sNome = ComboOpcoes.Text
'
'    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
'    If lErro <> SUCESSO Then gError 85185
'
'    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
'    If lErro <> SUCESSO Then gError 85186
'
'    Call BotaoLimpar_Click
'
'    Exit Sub
'
'Erro_BotaoGravar_Click:
'
'    Select Case gErr
'
'        Case 85183
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
'            ComboOpcoes.SetFocus
'
'        Case 85184, 85185, 85186
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172581)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ProdutoFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ProdutoFinal_Validate
'
'    giProdInicial = 0
'
'    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
'    If lErro <> SUCESSO And lErro <> 27095 Then gError 85187
'
'    If lErro <> SUCESSO Then gError 85188
'
'    Exit Sub
'
'Erro_ProdutoFinal_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 85187
'
'        Case 85188
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172582)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ProdutoInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ProdutoInicial_Validate
'
'    giProdInicial = 1
'
'    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
'    If lErro <> SUCESSO And lErro <> 27095 Then gError 85189
'
'    If lErro <> SUCESSO Then gError 85190
'
'    Exit Sub
'
'Erro_ProdutoInicial_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 85189
'
'        Case 85190
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172583)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
''monta a express�o de sele��o de relat�rio
'
'Dim sExpressao As String
'Dim lErro As Long
'
'On Error GoTo Erro_Monta_Expressao_Selecao
'
'   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)
'
'   If sProd_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)
'
'    End If
'
'
'
'     If TipoInicial.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "TipoProduto  >= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoInicial.Text)))
'
'    End If
'
'    If TipoFinal.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "TipoProduto <= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoFinal.Text)))
'
'    End If
'
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If
'
'    Monta_Expressao_Selecao = SUCESSO
'
'    Exit Function
'
'Erro_Monta_Expressao_Selecao:
'
'    Monta_Expressao_Selecao = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172584)
'
'    End Select
'
'    Exit Function
'
'End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ANALISE_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Estoque de MP Dispon�veis"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpRelESTMPDispOCX"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

'        If Me.ActiveControl Is ProdutoInicial Then
'            Call LabelProdutoDe_Click
'        ElseIf Me.ActiveControl Is ProdutoFinal Then
'            Call LabelProdutoAte_Click
'        End If

    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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
'
'Private Sub TipoInicial_Validate(Cancel As Boolean)
''Se mudar o tipo trazer dele os defaults para os campos da tela
'
'Dim lErro As Long
'Dim objTipoProduto As New ClassTipoDeProduto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoInicial_Validate
'
'    If Len(Trim(TipoInicial.Text)) <> 0 Then
'
'        'Critica o valor
'        lErro = Inteiro_Critica(Codigo_Extrai(TipoInicial.Text))
'        If lErro <> SUCESSO Then gError 85191
'
'        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoInicial.Text))
'
'        'L� o tipo
'        lErro = CF("TipoDeProduto_Le", objTipoProduto)
'        If lErro <> SUCESSO And lErro <> 22531 Then gError 85192
'
'        'Se n�o encontrar --> Erro
'        If lErro = 22531 Then gError 85193
'
'        TipoInicial.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
'
'    End If
'
'    Exit Sub
'
'Erro_TipoInicial_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 85191, 85192
'
'        Case 85193
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172585)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TipoFinal_Validate(Cancel As Boolean)
''Se mudar o tipo trazer dele os defaults para os campos da tela
'
'Dim lErro As Long
'Dim objTipoProduto As New ClassTipoDeProduto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoFinal_Validate
'
'    If Len(Trim(TipoFinal.Text)) <> 0 Then
'
'        'Critica o valor
'        lErro = Inteiro_Critica(Codigo_Extrai(TipoFinal.Text))
'        If lErro <> SUCESSO Then gError 85194
'
'        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoFinal.Text))
'
'        'L� o tipo
'        lErro = CF("TipoDeProduto_Le", objTipoProduto)
'        If lErro <> SUCESSO And lErro <> 22531 Then gError 85195
'
'        'Se n�o encontrar --> Erro
'        If lErro = 22531 Then gError 85196
'
'        TipoFinal.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
'
'    End If
'
'    Exit Sub
'
'Erro_TipoFinal_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 85194, 85195
'
'        Case 85196
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172586)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LabelTipoInicial_Click()
'
'Dim lErro As Long
'Dim objTipoProduto As ClassTipoDeProduto
'Dim colSelecao As Collection
'
'On Error GoTo Erro_LabelTipoInicial_Click
'
'    If Len(Trim(TipoInicial.Text)) <> 0 Then
'
'        Set objTipoProduto = New ClassTipoDeProduto
'        objTipoProduto.iTipo = Codigo_Extrai(TipoInicial.Text)
'
'
'    End If
'
'    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoInicial)
'
'    Exit Sub
'
'Erro_LabelTipoInicial_Click:
'
'    Select Case gErr
'
'         Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172587)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'Private Sub objEventoTipoInicial_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objTipoProduto As New ClassTipoDeProduto
'
'On Error GoTo Erro_objEventoTipoInicial_evSelecao
'
'    Set objTipoProduto = obj1
'
'    TipoInicial.Text = objTipoProduto.iTipo
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoTipoInicial_evSelecao:
'
'    Select Case Err
'
'       Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172588)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoTipoFinal_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objTipoProduto As New ClassTipoDeProduto
'
'On Error GoTo Erro_objEventoTipoFinal_evSelecao
'
'    Set objTipoProduto = obj1
'
'    TipoFinal.Text = objTipoProduto.iTipo
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoTipoFinal_evSelecao:
'
'    Select Case gErr
'
'       Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172589)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'
'Private Sub LabelTipoFinal_Click()
'
'Dim lErro As Long
'Dim colSelecao As Collection
'Dim objTipoProduto As ClassTipoDeProduto
'
'On Error GoTo Erro_LabelTipoFinal_Click
'
'    If Len(Trim(TipoFinal.Text)) <> 0 Then
'
'        Set objTipoProduto = New ClassTipoDeProduto
'        objTipoProduto.iTipo = Codigo_Extrai(TipoFinal.Text)
'
'    End If
'
'    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoFinal)
'
'    Exit Sub
'
'Erro_LabelTipoFinal_Click:
'
'    Select Case gErr
'
'         Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172590)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'