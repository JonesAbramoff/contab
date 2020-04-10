VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long, sProduto As String
Dim objProdutoFilial As New ClassProdutoFilial
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then

        objProdutoFilial.sProduto = objProduto.sCodigo
        
        'Traz os dados para a Tela
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then Error 33237
    
        'Coloca na tela o Produto selecionado
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
        
        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then Error 44972

        'Não encontrou o Produto
        If lErro = 25041 Then Error 44973
    
        'Preenche ProdutoDescricao com Descrição do Produto
        Descricao.Caption = objProduto.sDescricao
    
        'Preenche a Unidade de Medida
        LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
    
        objProdutoFilial.sProduto = objProduto.sCodigo
         
        lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque)
        If lErro <> SUCESSO And lErro <> 33253 Then Error 44974

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 33234, 33237, 44972, 44974

        Case 44973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155713)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get ClasseABC() As Object
     Set ClasseABC = objUserControl.Controls("ClasseABC")
End Property

Public Property Get IntRessup() As Object
     Set IntRessup = objUserControl.Controls("IntRessup")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get PontoPedido() As Object
     Set PontoPedido = objUserControl.Controls("PontoPedido")
End Property

Public Property Get EstoqueSeguranca() As Object
     Set EstoqueSeguranca = objUserControl.Controls("EstoqueSeguranca")
End Property

Public Property Get EstoqueMaximo() As Object
     Set EstoqueMaximo = objUserControl.Controls("EstoqueMaximo")
End Property

Public Property Get LoteEconomico() As Object
     Set LoteEconomico = objUserControl.Controls("LoteEconomico")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get ControleEstoque() As Object
     Set ControleEstoque = objUserControl.Controls("ControleEstoque")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get LblUMEstoque() As Object
     Set LblUMEstoque = objUserControl.Controls("LblUMEstoque")
End Property

Public Property Get TvwProduto() As Object
     Set TvwProduto = objUserControl.Controls("TvwProduto")
End Property

Public Property Get LabelProduto() As Object
     Set LabelProduto = objUserControl.Controls("LabelProduto")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 33242

    'Limpa a Tela
    Call Limpa_Tela_Estoque
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 33242

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 155714)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 33243

    'Limpa a Tela
    Call Limpa_Tela_Estoque

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 33243

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155715)

    End Select

    Exit Sub

End Sub

Public Sub ClasseABC_Change()

    ClasseABC.Text = UCase(ClasseABC.Text)
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ClasseABC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sClasse As String

On Error GoTo Erro_ClasseABC_Validate

    If Len(Trim(ClasseABC.Text)) = 0 Then Exit Sub
    
    sClasse = ClasseABC.Text
    
    'Verifica se Classe é do Tipo A, B ou C
    If sClasse <> "A" Then
        
        If sClasse <> "B" Then
        
            If sClasse <> "C" Then Error 33271
            
        End If
        
    End If
    
    Exit Sub

Erro_ClasseABC_Validate:

    Cancel = True


    Select Case Err

        Case 33271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSE_PRODUTO_INEXISTENTE", Err, ClasseABC.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155716)

    End Select

    Exit Sub

End Sub

Public Sub ControleEstoque_Click()

Dim lErro As Long

On Error GoTo Error_ControleEstoque_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se foi preenchida a ComboBox ControleEstoque
    If ControleEstoque.ListIndex = -1 Then Exit Sub

    'Verifica o tipo de controle de estoque
    If ControleEstoque.ItemData(ControleEstoque.ListIndex) = PRODUTO_CONTROLE_SEM_ESTOQUE Then
    
        'Limpa e desabilita os campos estoque de seguranca, est max, ponto de pedido, lote economico e intervalo de ressup,
        EstoqueSeguranca.Text = ""
        EstoqueSeguranca.Enabled = False
        EstoqueMaximo.Text = ""
        EstoqueMaximo.Enabled = False
        PontoPedido.Text = ""
        PontoPedido.Enabled = False
        LoteEconomico.Text = ""
        LoteEconomico.Enabled = False
    Else
        EstoqueSeguranca.Enabled = True
        EstoqueMaximo.Enabled = True
        PontoPedido.Enabled = True
        LoteEconomico.Enabled = True
    End If

    Exit Sub

Error_ControleEstoque_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155717)

    End Select

    Exit Sub

End Sub


Public Sub EstoqueMaximo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub EstoqueMaximo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sQuantidade As String

On Error GoTo Erro_EstoqueMaximo_Validate
        
    'Verifica se EstoqueMáximo foi digitado
    If Len(Trim(EstoqueMaximo.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(EstoqueMaximo.Text)
    If lErro <> SUCESSO Then Error 33273
    
    Exit Sub

Erro_EstoqueMaximo_Validate:

    Cancel = True


    Select Case Err

        Case 33273

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155718)

    End Select

    Exit Sub

End Sub

Public Sub EstoqueSeguranca_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub EstoqueSeguranca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstoqueSeguranca_Validate

    'Verifica se EstoqueSegurança foi digitado
    If Len(Trim(EstoqueSeguranca.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(EstoqueSeguranca.Text)
    If lErro <> SUCESSO Then Error 33272
    
    Exit Sub

Erro_EstoqueSeguranca_Validate:

    Cancel = True


    Select Case Err

        Case 33272

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155719)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento

    'Carrega a árvore de Produtos com os Produtos do BD
    lErro = CF("Carga_Arvore_Produto", TvwProduto.Nodes)
    If lErro <> SUCESSO Then Error 33235

    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 33236

    'Se módulo de Estoque não faz parte do pacote
    If gcolModulo.Ativo(MODULO_ESTOQUE) <> MODULO_ATIVO Then
    
        'Se o estoque não estiver ativo não permitir selecionar controle de estoque na combo: nem reserva+estoque nem estoque
        'Selecionar a outra opção e desabilitar a combo
        ControleEstoque.ListIndex = 2
        ControleEstoque.Enabled = False

    End If
    
    PontoPedido.Format = FORMATO_ESTOQUE
    LoteEconomico.Format = FORMATO_ESTOQUE
    EstoqueMaximo.Format = FORMATO_ESTOQUE
    EstoqueSeguranca.Format = FORMATO_ESTOQUE
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 33235, 33236
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155720)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
End Sub

Public Sub IntRessup_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(IntRessup, iAlterado)
    
End Sub

Public Sub LoteEconomico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub LoteEconomico_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteEconomico_Validate

    'Verifica se LoteEconômico foi digitado
    If Len(Trim(LoteEconomico.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(LoteEconomico.Text)
    If lErro <> SUCESSO Then Error 33275
    
    Exit Sub

Erro_LoteEconomico_Validate:

    Cancel = True


    Select Case Err

        Case 33275

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155721)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 33237

    'Coloca na tela o Produto selecionado
    
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
    Call Produto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 46583

        Case 33237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155722)

    End Select

    Exit Sub

End Sub

Public Sub PontoPedido_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PontoPedido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PontoPedido_Validate

    'Verifica se PontoPedido foi digitado
    If Len(Trim(PontoPedido.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PontoPedido.Text)
    If lErro <> SUCESSO Then Error 33274
    
    Exit Sub

Erro_PontoPedido_Validate:

    Cancel = True


    Select Case Err

        Case 33274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155723)

    End Select

    Exit Sub

End Sub

Public Sub Produto_Change()

    iProdutoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim iCodigo As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = 1 Then

        'Verifica preenchimento de Produto
        If Len(Trim(Produto.ClipText)) > 0 Then

            sProduto = Produto.Text

            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 Then Error 33240

            'Não encontrou o Produto
            If lErro = 25041 Then Error 33241

            'Preenche ProdutoDescricao com Descrição do Produto
            Descricao.Caption = objProduto.sDescricao

            'Preenche a Unidade de Medida
            LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
       
            objProdutoFilial.sProduto = objProduto.sCodigo
            
            lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque)
            If lErro <> SUCESSO And lErro <> 33253 Then Error 46585
        
        Else
            Descricao.Caption = ""
            LblUMEstoque.Caption = ""

        End If

        iProdutoAlterado = 0

    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True


    Select Case Err

        Case 33240
        
        Case 33241
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                'Segura o foco

            End If
            
        Case 46585
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155724)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

    'Chama a tela de browse se Produto
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

End Sub

Public Sub TvwProduto_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwProduto_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Produto_Netos", objNode, TvwProduto.Nodes)
        If lErro <> SUCESSO Then Error 48087
        
    End If
    
    Exit Sub
    
Erro_TvwProduto_Expand:

    Select Case Err
    
        Case 48087
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155725)
        
    End Select
        
    Exit Sub
    
End Sub

Public Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim sCodigo As String
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_TvwProduto_NodeClick

    'Armazena key do nó clicado sem caracter inicial
    sCodigo = right(Node.Key, Len(Node.Key) - 1)

    'Verifica se produto tem filhos
    If Node.Children > 0 Then Exit Sub

    objProduto.sCodigo = sCodigo
    'Lê Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 33238

    'Verifica se Produto é gerencial
    If objProduto.iGerencial = GERENCIAL Then Exit Sub

    'Mostra Unidade de Medida na tela
    LblUMEstoque.Caption = objProduto.sSiglaUMEstoque

    lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then Error 33239

    Call Produto_Validate(bSGECancelDummy)

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_TvwProduto_NodeClick:

    Select Case Err

        Case 33238, 33239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155726)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_Estoque() As Long
'Limpa os campos tela Estoque

Dim lErro As Long

    'Função generica que limpa campos da tela
    Call Limpa_Tela(Me)

    Descricao.Caption = ""
    LblUMEstoque.Caption = ""
    ClasseABC.Text = ""

    ControleEstoque.ListIndex = -1

    iAlterado = 0

End Function

Function Gravar_Registro() As Long
'Verifica se dados de Estoque necessários foram preenchidos

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim iControleEstoque As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 33244

    'Verifica se foi preenchida o Controle Estoque
    If Len(Trim(ControleEstoque.Text)) = 0 Then Error 33245

    'Lê os dados da Tela relacionados ao Estoque
    lErro = Move_Tela_Memoria(objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then Error 33246

    'Grava o Cliente no BD
    lErro = CF("Estoque_Grava", objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then Error 33247

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = Err

    Select Case Err

        Case 33244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 33245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTROLE_ESTOQUE_NAO_PREENCHIDO", Err)

        Case 33246, 33247

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155727)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objProdutoFilial.iFilialEmpresa = giFilialEmpresa

    'Verifica se o Produto foi preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 33248

        'testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then objProdutoFilial.sProduto = sProduto

    End If

    'Verifica os demais dados
    
    'Move o que esta selecionado na Combobox ControleEstoque para iControleEstoque
    If ControleEstoque.ListIndex <> -1 Then iControleEstoque = ControleEstoque.ItemData(ControleEstoque.ListIndex)

    If Len(Trim(EstoqueSeguranca.Text)) > 0 Then
         objProdutoFilial.dEstoqueSeguranca = CDbl(EstoqueSeguranca.Text)
    End If
    
    If Len(Trim(EstoqueMaximo.Text)) > 0 Then
        objProdutoFilial.dEstoqueMaximo = CDbl(EstoqueMaximo.Text)
    End If
    
    If Len(Trim(PontoPedido.Text)) > 0 Then
        objProdutoFilial.dPontoPedido = CDbl(PontoPedido.Text)
    End If
    
    If Len(Trim(LoteEconomico.Text)) > 0 Then
        objProdutoFilial.dLoteEconomico = CDbl(LoteEconomico.Text)
    End If
    
    If Len(Trim(IntRessup.Text)) > 0 Then
        objProdutoFilial.iIntRessup = CInt(IntRessup.Text)
    End If

    objProdutoFilial.sClasseABC = ClasseABC.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 33248

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155728)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim iControleEstoque As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProdutosFilial"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then Error 33249

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objProdutoFilial.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "EstoqueSeguranca", objProdutoFilial.dEstoqueSeguranca, 0, "EstoqueSeguranca"
    colCampoValor.Add "EstoqueMaximo", objProdutoFilial.dEstoqueMaximo, 0, "EstoqueMaximo"
    colCampoValor.Add "PontoPedido", objProdutoFilial.dPontoPedido, 0, "PontoPedido"
    colCampoValor.Add "LoteEconomico", objProdutoFilial.dLoteEconomico, 0, "LoteEconomico"
    colCampoValor.Add "IntRessup", objProdutoFilial.iIntRessup, 0, "IntRessup"
    colCampoValor.Add "ClasseABC", objProdutoFilial.sClasseABC, STRING_PRODUTOFILIAL_CLASSEABC, "ClassABC"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objProdutoFilial.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 33249

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155729)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Preenche

    objProdutoFilial.sProduto = colCampoValor.Item("Produto").vValor

    If objProdutoFilial.sProduto <> "" Then

        objProduto.sCodigo = objProdutoFilial.sProduto

        lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
        If lErro <> SUCESSO Then Error 46586
        
        Call Produto_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 33250, 46586

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155730)

    End Select

    Exit Sub

End Sub

Private Function Traz_Estoque_Tela(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Traz os dados do Estoque para tela

Dim lErro As Long
Dim sProduto As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_Estoque_Tela
    
    'Coloca os demais dados do Estoque na tela
    If iControleEstoque > 0 Then
        ControleEstoque.ListIndex = iControleEstoque - 1
    Else
        ControleEstoque.ListIndex = -1
    End If
    
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    
    'Lê o ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then Error 33252

    'Não encontrou o ProdutoFilial ==> erro
    If lErro = 28261 Then Error 33253


    If objProdutoFilial.dEstoqueSeguranca <> 0 Then
        EstoqueSeguranca.Text = CStr(objProdutoFilial.dEstoqueSeguranca)
    Else
        EstoqueSeguranca.Text = ""
    End If

    If objProdutoFilial.dEstoqueMaximo <> 0 Then
        EstoqueMaximo.Text = CStr(objProdutoFilial.dEstoqueMaximo)
    Else
        EstoqueMaximo.Text = ""
    End If

    If objProdutoFilial.dPontoPedido <> 0 Then
        PontoPedido.Text = CStr(objProdutoFilial.dPontoPedido)
    Else
        PontoPedido.Text = ""
    End If

    If objProdutoFilial.dLoteEconomico <> 0 Then
        LoteEconomico.Text = CStr(objProdutoFilial.dLoteEconomico)
    Else
        LoteEconomico.Text = ""
    End If

    If objProdutoFilial.iIntRessup <> 0 Then
        IntRessup.Text = CStr(objProdutoFilial.iIntRessup)
    Else
        IntRessup.Text = ""
    End If

    ClasseABC.Text = objProdutoFilial.sClasseABC
    
    iAlterado = 0

    Traz_Estoque_Tela = SUCESSO

    Exit Function

Erro_Traz_Estoque_Tela:

    Traz_Estoque_Tela = Err

    Select Case Err

        Case 33251, 33252, 33254

        Case 33253

        Case 33255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155731)

    End Select

    Exit Function

End Function

'TRATA_PARAMETROS
'''lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque)
'''If lErro <> SUCESSO And lErro <> 33253 Then Error 33234
'''
'''If lErro <> SUCESSO Then

'OBJEVENTOPRODUTO_EVSELECAO
'''objProdutoFilial.sProduto = objProduto.sCodigo
'''
'''lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque)
'''If lErro <> SUCESSO And lErro <> 33253 Then Error 46583
    
'TVWPRODUTO_NODECLICK
'''objProdutoFilial.sProduto = objProduto.sCodigo
'''
'''lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque)
'''If lErro <> SUCESSO And lErro <> 33253 Then Error 46584



'TRAZ_ESTOQUE_TELA
'''lErro = Mascara_RetornaProdutoEnxuto(objProdutoFilial.sProduto, sProduto)
'''If lErro <> SUCESSO Then Error 33251
'''
''''Coloca o Produto na tela
'''Produto.PromptInclude = False
'''Produto.Text = sProduto
'''Produto.PromptInclude = True
'''Produto_Validate


Private Sub Limpa_Campos_ProdutoFilial()

    PontoPedido.Text = ""
    EstoqueMaximo.Text = ""
    EstoqueSeguranca.Text = ""
    LoteEconomico.Text = ""
    IntRessup.Text = ""
    ClasseABC.Text = ""
    EstoqueSeguranca.Enabled = True
    EstoqueMaximo.Enabled = True
    PontoPedido.Enabled = True
    LoteEconomico.Enabled = True

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Controle de Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Estoque"
    
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

'**** fim do trecho a ser copiado *****

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        End If
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

