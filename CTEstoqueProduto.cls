VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTEstoqueProduto"
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

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Constantes da Tela
Const ALMOXARIFADOS_DA_EMPRESA = -1
Const ALMOXARIFADOS_DA_FILIAL = 0

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim iLocalizacaoAlterada As Integer
Dim iFrameAtual As Integer
Dim iTelaVisivel As Integer

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Saldos = 2
Private Const TAB_Terceiros = 3

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Padrao() As Object
     Set Padrao = objUserControl.Controls("Padrao")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

Public Property Get LocalizacaoFisica() As Object
     Set LocalizacaoFisica = objUserControl.Controls("LocalizacaoFisica")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

'Public Property Get TvwProduto() As Object
'     Set TvwProduto = objUserControl.Controls("TvwProduto")
'End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get UnidMed() As Object
     Set UnidMed = objUserControl.Controls("UnidMed")
End Property

'Public Property Get LabelProduto() As Object
'     Set LabelProduto = objUserControl.Controls("LabelProduto")
'End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get QuantOutras3() As Object
     Set QuantOutras3 = objUserControl.Controls("QuantOutras3")
End Property

Public Property Get QuantDemo3() As Object
     Set QuantDemo3 = objUserControl.Controls("QuantDemo3")
End Property

Public Property Get QuantConsig3() As Object
     Set QuantConsig3 = objUserControl.Controls("QuantConsig3")
End Property

Public Property Get QuantConserto3() As Object
     Set QuantConserto3 = objUserControl.Controls("QuantConserto3")
End Property

Public Property Get ValorOutras3() As Object
     Set ValorOutras3 = objUserControl.Controls("ValorOutras3")
End Property

Public Property Get ValorDemo3() As Object
     Set ValorDemo3 = objUserControl.Controls("ValorDemo3")
End Property

Public Property Get ValorConsig3() As Object
     Set ValorConsig3 = objUserControl.Controls("ValorConsig3")
End Property

Public Property Get ValorConserto3() As Object
     Set ValorConserto3 = objUserControl.Controls("ValorConserto3")
End Property

Public Property Get ValorBenef3() As Object
     Set ValorBenef3 = objUserControl.Controls("ValorBenef3")
End Property

Public Property Get QuantBenef3() As Object
     Set QuantBenef3 = objUserControl.Controls("QuantBenef3")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get QuantOutras() As Object
     Set QuantOutras = objUserControl.Controls("QuantOutras")
End Property

Public Property Get QuantDemo() As Object
     Set QuantDemo = objUserControl.Controls("QuantDemo")
End Property

Public Property Get QuantConsig() As Object
     Set QuantConsig = objUserControl.Controls("QuantConsig")
End Property

Public Property Get QuantConserto() As Object
     Set QuantConserto = objUserControl.Controls("QuantConserto")
End Property

Public Property Get ValorOutras() As Object
     Set ValorOutras = objUserControl.Controls("ValorOutras")
End Property

Public Property Get ValorDemo() As Object
     Set ValorDemo = objUserControl.Controls("ValorDemo")
End Property

Public Property Get ValorConsig() As Object
     Set ValorConsig = objUserControl.Controls("ValorConsig")
End Property

Public Property Get ValorConserto() As Object
     Set ValorConserto = objUserControl.Controls("ValorConserto")
End Property

Public Property Get ValorBenef() As Object
     Set ValorBenef = objUserControl.Controls("ValorBenef")
End Property

Public Property Get QuantBenef() As Object
     Set QuantBenef = objUserControl.Controls("QuantBenef")
End Property

Public Property Get UnidMedLabel() As Object
     Set UnidMedLabel = objUserControl.Controls("UnidMedLabel")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get AlmoxarifadoLabel() As Object
     Set AlmoxarifadoLabel = objUserControl.Controls("AlmoxarifadoLabel")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get QuantPedido() As Object
     Set QuantPedido = objUserControl.Controls("QuantPedido")
End Property

Public Property Get LabelQuantReqCompra() As Object
     Set LabelQuantReqCompra = objUserControl.Controls("LabelQuantReqCompra")
End Property

Public Property Get LabelValReqCompra() As Object
     Set LabelValReqCompra = objUserControl.Controls("LabelValReqCompra")
End Property

Public Property Get LabelQuantPedCompra() As Object
     Set LabelQuantPedCompra = objUserControl.Controls("LabelQuantPedCompra")
End Property

Public Property Get LabelValPedCompra() As Object
     Set LabelValPedCompra = objUserControl.Controls("LabelValPedCompra")
End Property

Public Property Get QuantPedidoLabel() As Object
     Set QuantPedidoLabel = objUserControl.Controls("QuantPedidoLabel")
End Property

Public Property Get ValorPedido() As Object
     Set ValorPedido = objUserControl.Controls("ValorPedido")
End Property

Public Property Get QuantInd() As Object
     Set QuantInd = objUserControl.Controls("QuantInd")
End Property

Public Property Get QuantDefeituosa() As Object
     Set QuantDefeituosa = objUserControl.Controls("QuantDefeituosa")
End Property

Public Property Get QuantEmpenhada() As Object
     Set QuantEmpenhada = objUserControl.Controls("QuantEmpenhada")
End Property

Public Property Get QuantReservada() As Object
     Set QuantReservada = objUserControl.Controls("QuantReservada")
End Property

Public Property Get QuantDispTotal() As Object
     Set QuantDispTotal = objUserControl.Controls("QuantDispTotal")
End Property

Public Property Get QuantTotal() As Object
     Set QuantTotal = objUserControl.Controls("QuantTotal")
End Property

Public Property Get ValorInd() As Object
     Set ValorInd = objUserControl.Controls("ValorInd")
End Property

Public Property Get ValorDefeituosa() As Object
     Set ValorDefeituosa = objUserControl.Controls("ValorDefeituosa")
End Property

Public Property Get ValorEmpenhada() As Object
     Set ValorEmpenhada = objUserControl.Controls("ValorEmpenhada")
End Property

Public Property Get ValorReservada() As Object
     Set ValorReservada = objUserControl.Controls("ValorReservada")
End Property

Public Property Get ValorDispTotal() As Object
     Set ValorDispTotal = objUserControl.Controls("ValorDispTotal")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get ValorOP() As Object
     Set ValorOP = objUserControl.Controls("ValorOP")
End Property

Public Property Get QuantOP() As Object
     Set QuantOP = objUserControl.Controls("QuantOP")
End Property

Public Property Get DataInventario() As Object
     Set DataInventario = objUserControl.Controls("DataInventario")
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

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get CustoLabel() As Object
    Set CustoLabel = objUserControl.Controls("CustoLabel")
End Property

Public Property Get CustoOutras3() As Object
     Set CustoOutras3 = objUserControl.Controls("CustoOutras3")
End Property

Public Property Get CustoDemo3() As Object
     Set CustoDemo3 = objUserControl.Controls("CustoDemo3")
End Property

Public Property Get CustoConsig3() As Object
     Set CustoConsig3 = objUserControl.Controls("CustoConsig3")
End Property

Public Property Get CustoConserto3() As Object
     Set CustoConserto3 = objUserControl.Controls("CustoConserto3")
End Property

Public Property Get CustoBenef3() As Object
     Set CustoBenef3 = objUserControl.Controls("CustoBenef3")
End Property

Public Property Get CustoOutras() As Object
     Set CustoOutras = objUserControl.Controls("CustoOutras")
End Property

Public Property Get CustoDemo() As Object
     Set CustoDemo = objUserControl.Controls("CustoDemo")
End Property

Public Property Get CustoConsig() As Object
     Set CustoConsig = objUserControl.Controls("CustoConsig")
End Property

Public Property Get CustoConserto() As Object
     Set CustoConserto = objUserControl.Controls("CustoConserto")
End Property

Public Property Get CustoBenef() As Object
     Set CustoBenef = objUserControl.Controls("CustoBenef")
End Property

Public Property Get ProdutoLabel1() As Object
     Set ProdutoLabel1 = objUserControl.Controls("ProdutoLabel1")
End Property
'--- fim dos properties get dos controles da tela

Public Sub ProdutoLabel1_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLabel1_Click

    sSelecao = "ControleEstoque <> ? AND FilialEmpresa=?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    colSelecao.Add giFilialEmpresa
    
    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83060
        
        objProduto.sCodigo = sProdutoFormatado

    End If
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_ProdutoLabel1_Click:

    Select Case gErr

        Case 83060

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155788)

    End Select

    Exit Sub

End Sub

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Almoxarifado_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim iCodigo As Integer
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Almoxarifado_Click

    'Verifica preenchimento de Almoxarifado
    If Almoxarifado.ListIndex <> -1 And iTelaVisivel = True Then

        'Verifica preenchimento de Produto
        If Len(Trim(Produto.ClipText)) > 0 Then

            sProduto = Produto.Text

            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
            If lErro <> SUCESSO And lErro <> 25041 Then gError 28430

            If lErro = 25041 Then gError 28431

            iCodigo = Codigo_Extrai(Almoxarifado.Text)

            lErro = EstoqueProduto_TrazTela(objProduto, iCodigo)
            If lErro <> SUCESSO Then gError 20771

            'Preenche objProdutoFilial
            objProdutoFilial.iFilialEmpresa = giFilialEmpresa
            objProdutoFilial.sProduto = objProduto.sCodigo
            
            'Lê ProdutoFilial
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO Then gError 28449

            If objProdutoFilial.iAlmoxarifado = iCodigo Then
                Padrao.Value = MARCADO
            Else
                Padrao.Value = DESMARCADO
            End If
            
        End If
        
        For iIndice = 0 To 1
            AlmoxarifadoLabel(iIndice).Caption = Almoxarifado.Text
        Next
        
        iAlterado = 0

    End If

    Exit Sub

Erro_Almoxarifado_Click:

    Select Case gErr

        Case 20771, 28449

        Case 28430
            Produto.SetFocus

        Case 28431
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Produto.SetFocus
            End If

            Descricao.Caption = ""
            UnidMed.Caption = ""
            Call Limpa_Campos_EstoqueProduto1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155789)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 28408

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 28409

    'Não encontrou o Produto
    If lErro = 25041 Then gError 28410

    If Almoxarifado.ItemData(Almoxarifado.ListIndex) = 0 Or Almoxarifado.ItemData(Almoxarifado.ListIndex) = -1 Then gError 57774
    
    'Verifica preenchimento de Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then gError 28411

    objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)
    
    'Lê Almoxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 28412

    'Não encontrou o Almoxarifado
    If lErro = 25056 Then gError 28413

    'Preenche objEstoqueProduto
    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    objEstoqueProduto.iFilialEmpresa = objAlmoxarifado.iFilialEmpresa
    
    'Lê o EstoqueProduto
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 28414

    'Não encontrou o Estoque Produto
    If lErro = 21306 Then gError 28415

    'Pede confirmação para exclusão Estoque Produto
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ESTOQUE_PRODUTO", objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

    If vbMsgRes = vbYes Then

        objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("EstoqueProduto_Exclui", objEstoqueProduto)
        If lErro <> SUCESSO Then gError 28416

        'Limpa a tela
        Call Limpa_Tela_EstoqueProduto

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 28408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 28409, 28412, 28414, 28416

        Case 28410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 28411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)

        Case 28413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)

        Case 28415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 57774
             lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_ALMOXARIFADO", gErr)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155790)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 28375

    Call Limpa_Tela_EstoqueProduto

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

       Case 28375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155791)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 28376

    'Limpa a tela
    Call Limpa_Tela_EstoqueProduto

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 28376

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155792)

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
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iTelaVisivel = False
    
    Set objEventoContaContabil = New AdmEvento
    Set objEventoProduto = New AdmEvento

    'Inicializa a máscara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 28365
    
    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 22159

    ContaContabil.Mask = sMascaraConta

'    'Inicializa a Lista de Produtos
'    lErro = CF("Carga_Arvore_Produto_Inventariado",TvwProduto.Nodes)
'    If lErro <> SUCESSO And lErro <> 52152 Then gError 28367
    
    If lErro = 52152 Then gError 52153
    
    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 22159, 28365, 28367
        
        Case 52153
            lErro = Rotina_Erro(vbOKOnly, "PRODUTO_INVENTARIADO_INEXISTENTE", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155793)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Limpa_Tela_EstoqueProduto() As Long
'Limpa a Tela EstoqueProduto

Dim lErro As Long
Dim iIndice As Integer

    'Funcção genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    'Limpa os campos da tela que não foram limpos pela função acima
    Padrao.Value = 0
    Descricao.Caption = ""
    UnidMed.Caption = ""
    Almoxarifado.Clear
    
    For iIndice = 0 To 1
        ProdutoLabel(iIndice).Caption = ""
    Next
    
    For iIndice = 0 To 1
        UnidMedLabel(iIndice).Caption = ""
    Next
    
    For iIndice = 0 To 1
        AlmoxarifadoLabel(iIndice).Caption = ""
    Next
    
    QuantDispTotal.Caption = ""
    QuantInd.Caption = ""
    QuantReservada.Caption = ""
    QuantDefeituosa.Caption = ""
    QuantTotal.Caption = ""
    QuantEmpenhada.Caption = ""
    QuantOP.Caption = ""
    QuantPedido.Caption = ""
    DataInventario.Caption = ""
    LabelQuantPedCompra.Caption = ""
    LabelQuantReqCompra.Caption = ""
    LabelValReqCompra.Caption = ""
    LabelValPedCompra.Caption = ""
    
    QuantConsig.Caption = ""
    QuantConsig3.Caption = ""
    QuantOutras.Caption = ""
    QuantOutras3.Caption = ""
    ValorDispTotal.Caption = ""
    ValorInd.Caption = ""
    ValorReservada.Caption = ""
    ValorDefeituosa.Caption = ""
    ValorTotal.Caption = ""
    ValorEmpenhada.Caption = ""
    ValorOP.Caption = ""
    ValorPedido.Caption = ""
    ValorConsig.Caption = ""
    ValorConsig3.Caption = ""
    ValorOutras.Caption = ""
    ValorOutras3.Caption = ""
    CustoConsig.Caption = ""
    CustoConsig3.Caption = ""
    CustoOutras.Caption = ""
    CustoOutras3.Caption = ""
    CustoLabel(1).Caption = ""
        
'    If giTipoVersao = VERSAO_FULL Then
        QuantConserto.Caption = ""
        QuantBenef.Caption = ""
        QuantBenef3.Caption = ""
        QuantConserto3.Caption = ""
        QuantDemo.Caption = ""
        QuantDemo3.Caption = ""
        QuantBenef.Caption = ""
        QuantBenef3.Caption = ""
        
        ValorConserto.Caption = ""
        ValorConserto3.Caption = ""
        ValorDemo.Caption = ""
        ValorDemo3.Caption = ""
        ValorBenef.Caption = ""
        ValorBenef3.Caption = ""
        CustoConserto.Caption = ""
        CustoConserto3.Caption = ""
        CustoDemo.Caption = ""
        CustoDemo3.Caption = ""
        CustoBenef.Caption = ""
        CustoBenef3.Caption = ""
'    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Limpa_Tela_EstoqueProduto = SUCESSO

    iLocalizacaoAlterada = 0
    
    iAlterado = 0

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoContaContabil = Nothing
    Set objEventoProduto = Nothing
   
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
   lErro = ComandoSeta_Liberar(Me.Name)
      
End Sub

Public Sub LocalizacaoFisica_Change()

    iLocalizacaoAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49545

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 49546

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 49547

        'conta não cadastrada
        If lErro = 5700 Then gError 49548

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 49545, 49547
    
        Case 49546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 49548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155794)

    End Select

    Exit Sub

End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 49539

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

        Case 49539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155795)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabil.Text = ""

    Else

        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 49538

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 49538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155796)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 83061

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 83062

    'Coloca na tela o Produto selecionado
    UnidMed.Caption = objProduto.sSiglaUMEstoque

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 83063
    
    For iIndice = 0 To 1
        ProdutoLabel(iIndice).Caption = Produto.Text
    Next

    Call Produto_Validate(bSGECancelDummy)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 83064

    Me.Show
    
    iAlterado = 0

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 83061, 83063, 83064

        Case 83062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155797)

    End Select

    Exit Sub

End Sub

Public Sub Padrao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iProdutoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate
   
    'função que preenche a tela e trata o produto no caso de ser ou nao inventariado
    lErro = Trata_Produto()
    If lErro <> SUCESSO Then gError 52157
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True


    Select Case gErr
    
        Case 52157

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155798)

    End Select

    Exit Sub

End Sub

Function Trata_Produto() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim colAlmoxarifados As New Collection
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Produto

    If iProdutoAlterado = 1 Then
    
        Call Limpa_Campos_EstoqueProduto1

        'Verifica preenchimento de Produto
        If Len(Trim(Produto.ClipText)) > 0 Then

            sProduto = Produto.Text

            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 28432
            
            'O Produto não está cadastrado
            If lErro = 25041 Then gError 28433

            'O Produto é Gerencial
            If lErro = 25043 Then gError 33365
            
            'O Produto não é Inventariado
            If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 52155
            
            'Preenche ProdutoDescricao com Descrição do Produto
            Descricao.Caption = objProduto.sDescricao

            'Preenche a Unidade de Medida
            UnidMed.Caption = objProduto.sSiglaUMEstoque
            
            For iIndice = 0 To 1
                
                UnidMedLabel(iIndice).Caption = objProduto.sSiglaUMEstoque
                ProdutoLabel(iIndice).Caption = Produto
            Next

            'Carrega a ComboBox de Almoxarifado ligados ao produto
            lErro = Carrega_Almoxarifados(objProduto.sCodigo, colAlmoxarifados)
            If lErro <> SUCESSO Then gError 28447
            
            If colAlmoxarifados.Count > 0 Then

                'Preenche objProdutoFilial
                objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                objProdutoFilial.sProduto = objProduto.sCodigo
                
                'Lê ProdutoFilial
                lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                If lErro <> SUCESSO And lErro <> 28261 Then gError 28436
    
                'Se achou
                If lErro = SUCESSO Then
                
                    'se tem um almoxarifado padrao. Seleciona-o
                    If objProdutoFilial.iAlmoxarifado <> 0 Then
                    
                        'Seleciona o almoxarifado padrao.
                        For iIndice = 0 To Almoxarifado.ListCount - 1
                        
                            If Almoxarifado.ItemData(iIndice) = objProdutoFilial.iAlmoxarifado Then
                                Almoxarifado.ListIndex = iIndice
                                For iIndice1 = 0 To 1
                                    AlmoxarifadoLabel(iIndice1).Caption = Almoxarifado.Text
                                Next
                                Exit For
                            End If
                            
                        Next
                        
                        Padrao.Value = MARCADO
                        
                    Else
                    
                        'se não houver um almoxarifado padrão ==> seleciona o primeiro da lista
                        Padrao.Value = DESMARCADO
                        
                        Almoxarifado.ListIndex = 0
                        
                    End If
                    
                Else
                
                    'se o produto não está associado a filial em questao
                    Call Limpa_Campos_EstoqueProduto1
                    
                End If
                
            Else
                
                gError 69842
                'se não há almoxarifados associados ao produto em questao
                Call Limpa_Campos_EstoqueProduto1
                
            End If
        
        'Se Produto não está preenchido
        Else
        
            'Limpa a tela
            Call Limpa_Tela_EstoqueProduto
            
        End If

        iProdutoAlterado = 0
        
        iAlterado = 0

    End If

    Trata_Produto = SUCESSO

    Exit Function
    
Erro_Trata_Produto:
    
    Trata_Produto = gErr
    
    Select Case gErr

        Case 28432, 33365
            Produto.SetFocus
            
        Case 28433
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Descricao.Caption = ""
                UnidMed.Caption = ""
                Call Limpa_Campos_EstoqueProduto1
                Produto.SetFocus
            End If

        Case 28436, 28447
        
        Case 52155
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
            Produto.SetFocus
        
        Case 69842
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO1", gErr, objProduto.sCodigo, giFilialEmpresa)
            Produto.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155799)

    End Select

    Exit Function

End Function

Private Function Trata_Produto1(objEstoqueProduto As ClassEstoqueProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim colAlmoxarifados As New Collection
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim iAchou As Integer

On Error GoTo Erro_Trata_Produto1

    'Limpa a ComboBox Almoxarifado
    Almoxarifado.Clear
    For iIndice = 0 To 1
        AlmoxarifadoLabel(iIndice).Caption = ""
    Next

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 55123
        
        'O Produto não está cadastrado
        If lErro = 25041 Then gError 55124
    
        'O Produto é Gerencial
        If lErro = 25043 Then gError 55125
        
        'O Produto não é Inventariado
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 55126
        
        'Preenche ProdutoDescricao com Descrição do Produto
        Descricao.Caption = objProduto.sDescricao

        'Preenche a Unidade de Medida
        UnidMed.Caption = objProduto.sSiglaUMEstoque
        
        For iIndice = 0 To 1
            
            UnidMedLabel(iIndice).Caption = objProduto.sSiglaUMEstoque
            ProdutoLabel(iIndice).Caption = Produto
        Next

        'Preenche objProdutoFilial
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        objProdutoFilial.sProduto = objProduto.sCodigo
        'Lê ProdutoFilial
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 55128
        
        'se não existir a associação do produto com a filial em questão
        If lErro = 28261 Then gError 55162
        
        'Carrega a ComboBox de Almoxarifado ligados ao produto
        lErro = Carrega_Almoxarifados(objProduto.sCodigo, colAlmoxarifados)
        If lErro <> SUCESSO Then gError 55127
        
        'se o almoxarifado não foi passado como parametro
        If objEstoqueProduto.iAlmoxarifado = 0 Then
        
            If colAlmoxarifados.Count > 0 Then
    
                'se houver almoxarifado padrao da filial
                If objProdutoFilial.iAlmoxarifado <> NAO_PADRAO Then
                    'Seleciona o almoxarifado padrao
                    For iIndice = 0 To Almoxarifado.ListCount - 1
                        If Almoxarifado.ItemData(iIndice) = objProdutoFilial.iAlmoxarifado Then
                            Almoxarifado.ListIndex = iIndice
                            
                            For iIndice1 = 0 To 1
                                AlmoxarifadoLabel(iIndice1).Caption = Almoxarifado.Text
                            Next
                            Exit For
                        End If
                    Next
                    Padrao.Value = MARCADO
                Else
                    Padrao.Value = DESMARCADO
                    Almoxarifado.ListIndex = 0
                End If
                
            Else
                Call Limpa_Campos_EstoqueProduto1
            End If

        Else
            'se o almoxarifado foi passado como parametro
            
            iAchou = 0
            'seleciona o almoxarifado passado como parametro
            For iIndice = 0 To Almoxarifado.ListCount - 1
                If Almoxarifado.ItemData(iIndice) = objEstoqueProduto.iAlmoxarifado Then
                    Almoxarifado.ListIndex = iIndice
                    For iIndice1 = 0 To 1
                        AlmoxarifadoLabel(iIndice1).Caption = Almoxarifado.Text
                    Next
                    
                    
                    iAchou = 1
                    If objEstoqueProduto.iAlmoxarifado = objProdutoFilial.iAlmoxarifado Then Padrao.Value = MARCADO
                End If
            Next
        
            If iAchou = 0 Then gError 55163
            
        End If
        
    'Se Produto não está preenchido
    Else
    
        'Limpa a tela
        Call Limpa_Campos_EstoqueProduto1
        
    End If

    If Almoxarifado.ListIndex <> -1 Then

        lErro = EstoqueProduto_TrazTela(objProduto, Almoxarifado.ItemData(Almoxarifado.ListIndex))
        If lErro <> SUCESSO Then gError 55164
                            
    End If

    Trata_Produto1 = SUCESSO

    Exit Function

Erro_Trata_Produto1:
    
    Trata_Produto1 = gErr
    
    Select Case gErr

        Case 55123, 55125, 55127, 55128, 55164
            
        Case 55124
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
            
        Case 55126
            Call Rotina_Erro(vbOKOnly, "PRODUTO_INVENTARIADO_INEXISTENTE2", gErr, Produto.Text)
            
        Case 55162
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, Produto.Text, giFilialEmpresa)
            
        Case 55163
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", gErr, objEstoqueProduto.iAlmoxarifado, Produto.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155800)

    End Select

    Exit Function

End Function

Public Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_PRODUTOS_ALMOXARIFADO_DADOS_PRINCIPAIS
                
            Case TAB_Saldos
                Parent.HelpContextID = IDH_PRODUTOS_ALMOXARIFADO_SALDOS
                        
            Case TAB_Terceiros
                Parent.HelpContextID = IDH_PRODUTOS_ALMOXARIFADO_SALDOS_TERCEIROS
                
        End Select

    End If

End Sub

'Public Sub TvwProduto_Expand(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If Node.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos_Inventariado",Node, TvwProduto.Nodes)
'        If lErro <> SUCESSO Then gError 52154
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case gErr
'
'        Case 52154
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155801)
'
'    End Select
'
'    Exit Sub
'End Sub
'
'Public Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
''??? deveria usar Traz_EstoqueProduto_Tela
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim sCodigo As String
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_TvwProduto_NodeClick
'
'    'Armazena key do nó clicado sem caracter inicial
'    sCodigo = Right(Node.Key, Len(Node.Key) - 1)
'
'    'Verifica se produto tem filhos
'    If Node.Children > 0 Then Exit Sub
'
'    objProduto.sCodigo = sCodigo
'    'Lê Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 28368
'
'    'Verifica se Produto é gerencial
'    If objProduto.iGerencial = GERENCIAL Then Exit Sub
'
'    'Mostra Unidade de Medida na tela
'    UnidMed.Caption = objProduto.sSiglaUMEstoque
'
'    lErro = CF("Traz_Produto_MaskEd",sCodigo, Produto, Descricao)
'    If lErro <> SUCESSO Then gError 28369
'
'    For iIndice = 0 To 1
'        ProdutoLabel(iIndice).Caption = Produto
'    Next
'
'    Call Produto_Validate(bSGECancelDummy)
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_TvwProduto_NodeClick:
'
'    Select Case gErr
'
'        Case 28368, 28369
'
'        Case 52013
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155802)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Function Trata_Parametros(Optional objEstoqueProduto As ClassEstoqueProduto) As Long

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_Trata_Parametros

    'Se há um EstoqueProduto selecionado, exibir seus dados
    If Not (objEstoqueProduto Is Nothing) Then

        sCodigo = objEstoqueProduto.sProduto
        lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, Descricao)
        If lErro <> SUCESSO Then gError 28370
                
        'funcao que preenche a tela e trata o produto no caso de ser ou nao inventariado
        lErro = Trata_Produto1(objEstoqueProduto)
        If lErro <> SUCESSO Then gError 52159
        
    End If
    
    iProdutoAlterado = 0
    iAlterado = 0
    iTelaVisivel = True

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 28370, 52159
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155803)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim iIndice As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_Tela_Extrai

    'tulio300102
    'alteracao para filtrar por filial
    'Informa tabela associada à Tela
    sTabela = "EstoqueProduto_Almoxarifado"

    'Lê os dados da Tela EstoqueProduto
    lErro = Move_Tela_Memoria(objEstoqueProduto)
    If lErro <> SUCESSO Then gError 28373

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objEstoqueProduto.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Almoxarifado", objEstoqueProduto.iAlmoxarifado, 0, "Almoxarifado"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 28373

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155804)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_Tela_Preenche

    objEstoqueProduto.sProduto = colCampoValor.Item("Produto").vValor
    objEstoqueProduto.iAlmoxarifado = colCampoValor.Item("Almoxarifado").vValor

    If (objEstoqueProduto.sProduto <> "") And (objEstoqueProduto.iAlmoxarifado <> 0) Then

        'Traz os dados para tela
        lErro = Traz_EstoqueProduto_Tela(objEstoqueProduto)
        If lErro <> SUCESSO Then gError 28374

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 28374

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155805)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iAlmoxarifadoPadraoFilial As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 28377

    'Verifica preenchimento de Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then gError 28378

    If Almoxarifado.ItemData(Almoxarifado.ListIndex) = 0 Or Almoxarifado.ItemData(Almoxarifado.ListIndex) = -1 Then gError 57773
    
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 28379

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objEstoqueProduto)
    If lErro <> SUCESSO Then gError 28382

    lErro = Trata_Alteracao(objEstoqueProduto, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
    If lErro <> SUCESSO Then Error 32283

    'Almoxarifado padrão
    If Padrao.Value <> NAO_PADRAO Then
        iAlmoxarifadoPadraoFilial = Codigo_Extrai(Almoxarifado.Text)
    Else
        iAlmoxarifadoPadraoFilial = NAO_PADRAO
    End If
        
    'Chama EstoqueProduto_AtualizaAlmoxarifado
    lErro = CF("EstoqueProduto_AtualizaAlmoxarifado", objEstoqueProduto, iAlmoxarifadoPadraoFilial)
    If lErro <> SUCESSO Then gError 28383

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 32283

        Case 28377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 28378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)

        Case 28379, 28382, 28383

        Case 57773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_ALMOXARIFADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155806)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objEstoqueProduto As ClassEstoqueProduto) As Long
'Move os dados da Tela para objEstoqueProduto

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 28384

        objEstoqueProduto.sProduto = sProdutoFormatado

    End If

    If Almoxarifado.Text <> "<Todos Almoxarifados>" Then
        objEstoqueProduto.iAlmoxarifado = Codigo_Extrai(Almoxarifado.Text)
    End If
    
    'Preenche objEstoqueProduto
    If Len(Trim(LocalizacaoFisica.Text)) > 0 Then objEstoqueProduto.sLocalizacaoFisica = LocalizacaoFisica.Text
    
    'Verifica se o modulo de contabilidade esta ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
            'Guarda a conta corrente
            lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 22161
            
            objEstoqueProduto.sContaContabil = sContaFormatada
            
        End If
        
    Else
        objEstoqueProduto.sContaContabil = ""
        
    End If
    
    objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 22161, 28384

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155807)

    End Select

    Exit Function

End Function

Private Function Traz_EstoqueProduto_Tela(objEstoqueProduto As ClassEstoqueProduto) As Long
'Traz os dados da EstoqueProduto para a Tela

Dim lErro As Long
Dim iIndice As Integer
Dim sCodigo As String

On Error GoTo Erro_Traz_EstoqueProduto_Tela

    sCodigo = objEstoqueProduto.sProduto

    lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 28388
    
    For iIndice = 0 To 1
        ProdutoLabel(iIndice).Caption = Produto
    Next

    Call Produto_Validate(bSGECancelDummy)
    
    For iIndice = 0 To Almoxarifado.ListCount - 1
    
        If Almoxarifado.ItemData(iIndice) = objEstoqueProduto.iAlmoxarifado Then
            Almoxarifado.ListIndex = iIndice
            Exit For
        End If
        
    Next
    
    iAlterado = 0

    Traz_EstoqueProduto_Tela = SUCESSO

    Exit Function

Erro_Traz_EstoqueProduto_Tela:

    Traz_EstoqueProduto_Tela = gErr

    Select Case gErr

        Case 28388

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155808)

    End Select

    Exit Function

End Function

Private Function Carrega_Almoxarifados(sProduto As String, colAlmoxarifados As Collection) As Long
'Carrega os Almoxarifados do Produto
    
Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Carrega_Almoxarifados

    'Limpa a ComboBox Almoxarifado
    Almoxarifado.Clear
    
    'Lê todos os almoxarifados do produto na filial corrente
    lErro = CF("Produto_Le_Almoxarifados_Filial", sProduto, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 33631
    
    Almoxarifado.AddItem ("<Almoxarifados da Empresa>")
    Almoxarifado.ItemData(Almoxarifado.NewIndex) = -1
    
'    If giTipoVersao = VERSAO_FULL Then
        Almoxarifado.AddItem ("<Almoxarifados da Filial>")
        Almoxarifado.ItemData(Almoxarifado.NewIndex) = 0
    
'    End If
    
    'Insere na combo CategoriaProdutoItem
    For Each objAlmoxarifado In colAlmoxarifados
    
        'Insere na combo
        Almoxarifado.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifado.ItemData(Almoxarifado.NewIndex) = objAlmoxarifado.iCodigo
    
    Next

    Carrega_Almoxarifados = SUCESSO

    Exit Function

Erro_Carrega_Almoxarifados:

    Carrega_Almoxarifados = gErr

    Select Case gErr

        Case 33631
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155809)

    End Select

    Exit Function

End Function

Private Function EstoqueProduto_TrazTela(objProduto As ClassProduto, iCodigo As Integer) As Long
'Traz p/tela dados (saldos,...) referentes ao par (Produto,Almoxarifado)


'????? Receber o obj

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objEstoqueProdutoSUM As New ClassEstoqueProduto
Dim dQuant1 As Double
Dim dQuant2 As Double
Dim dQuant3 As Double
Dim sContaMascarada As String
Dim dCusto As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double
Dim dQuantPedidaReqCompra As Double
Dim dQuantidadePedCompra As Double

On Error GoTo Erro_EstoqueProduto_TrazTela

    objProdutoFilial.sProduto = objProduto.sCodigo
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    
    'Para Ler a quantidade Pedida
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 61123
    
    If lErro = 28261 Then gError 61124
        
    objEstoqueProdutoSUM.sProduto = objProduto.sCodigo
    objEstoqueProdutoSUM.iAlmoxarifado = iCodigo

    'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProdutoSUM, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 64014 Then gError 20791
    
    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = iCodigo
    
    Select Case Almoxarifado.ItemData(Almoxarifado.ListIndex)
        
        Case ALMOXARIFADOS_DA_EMPRESA
        
            'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Empresa
            lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 52979 Then gError 52980
                
            'Calcula o Custo para a Empresa Toda
            lErro = CalculaValor_Empresa_Toda(objProduto.sCodigo, objEstoqueProdutoSUM, dCusto, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3)
            If lErro <> SUCESSO Then gError 69837
                
        Case ALMOXARIFADOS_DA_FILIAL

            'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
            lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 64014 Then gError 64015
        
            lErro = Calcula_Valor_Tela(objProduto.sCodigo, objEstoqueProdutoSUM, dCusto, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, giFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 69854 Then gError 69838
        
        Case Else
            
            'Lê o Estoque Produto
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 33630
        
            lErro = Calcula_Valor_Tela(objProduto.sCodigo, objEstoqueProdutoSUM, dCusto, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, giFilialEmpresa)
            If lErro <> SUCESSO Then gError 49943
            
    End Select
    
    If lErro = SUCESSO Then
    
        CustoLabel(1).Caption = Format(dCusto, FORMATO_CUSTO)
        
        LocalizacaoFisica.Text = objEstoqueProduto.sLocalizacaoFisica
                
        'preenche a conta contabil
        If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
            If objEstoqueProduto.sContaContabil <> "" Then
                sContaMascarada = String(STRING_CONTA, 0)
        
                lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then gError 49546
            Else
                sContaMascarada = ""
            End If
        
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaMascarada
            ContaContabil.PromptInclude = True
        
        End If
        
        ValorDispTotal.Caption = Format(objEstoqueProduto.dQuantDisponivel * dCusto, "standard")
        QuantDispTotal.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
        
        dQuant1 = objEstoqueProduto.dQuantInd
        dQuant2 = objEstoqueProduto.dQuantRecIndl
        dQuant3 = dQuant1 + dQuant2

        QuantInd.Caption = Formata_Estoque((objEstoqueProduto.dQuantInd + objEstoqueProduto.dQuantRecIndl))
        ValorInd.Caption = Format(dQuant3 * dCusto, "standard")
        
        QuantReservada.Caption = Formata_Estoque(objEstoqueProduto.dQuantReservada + objEstoqueProduto.dQuantReservadaConsig)
        ValorReservada.Caption = Format((objEstoqueProduto.dQuantReservada + objEstoqueProduto.dQuantReservadaConsig) * dCusto, "standard")
        
        QuantDefeituosa.Caption = Formata_Estoque(objEstoqueProduto.dQuantDefeituosa)
        ValorDefeituosa.Caption = Format(objEstoqueProduto.dQuantDefeituosa * dCusto, "standard")
        
        QuantEmpenhada.Caption = Formata_Estoque(objEstoqueProduto.dQuantEmpenhada)
        ValorEmpenhada.Caption = Format(objEstoqueProduto.dQuantEmpenhada * dCusto, "standard")
        
        QuantTotal.Caption = Formata_Estoque(StrParaDbl(QuantDispTotal.Caption) + StrParaDbl(QuantInd.Caption) + StrParaDbl(QuantDefeituosa.Caption) + StrParaDbl(QuantReservada.Caption))
        ValorTotal.Caption = Format(StrParaDbl(ValorDispTotal.Caption) + StrParaDbl(ValorInd.Caption) + StrParaDbl(ValorDefeituosa.Caption) + StrParaDbl(ValorReservada.Caption), "Standard")
        
        QuantOP.Caption = Formata_Estoque(objEstoqueProduto.dQuantOP)
        ValorOP.Caption = Format(objEstoqueProduto.dQuantOP * dCusto, "standard")
        
        QuantPedido.Caption = Formata_Estoque(objProdutoFilial.dQuantPedida)
        ValorPedido.Caption = Format(objProdutoFilial.dQuantPedida * dCusto, "standard")
        
        If objEstoqueProduto.dtDataInventario <> DATA_NULA And objEstoqueProduto.dtDataInventario <> 0 Then
            DataInventario.Caption = Format(objEstoqueProduto.dtDataInventario, "dd/mm/yyyy")
        Else
            DataInventario.Caption = ""
        End If
        
        If objEstoqueProduto.dQuantConsig = 0 Then
            dCustoConsig = 0
        End If
        QuantConsig.Caption = Formata_Estoque(objEstoqueProduto.dQuantConsig)
        CustoConsig.Caption = Format(dCustoConsig, FORMATO_CUSTO)
        ValorConsig.Caption = Format(objEstoqueProduto.dQuantConsig * dCustoConsig, "standard")
        
        If objEstoqueProduto.dQuantConsig3 = 0 Then
            dCustoConsig3 = 0
        End If
        QuantConsig3.Caption = Formata_Estoque(objEstoqueProduto.dQuantConsig3)
        CustoConsig3.Caption = Format(dCustoConsig3, FORMATO_CUSTO)
        ValorConsig3.Caption = Format(objEstoqueProduto.dQuantConsig3 * dCustoConsig3, "standard")
        
        If objEstoqueProduto.dQuantOutras = 0 Then
            dCustoOutras = 0
        End If
        QuantOutras.Caption = Formata_Estoque(objEstoqueProduto.dQuantOutras)
        CustoOutras.Caption = Format(dCustoOutras, FORMATO_CUSTO)
        ValorOutras.Caption = Format(objEstoqueProduto.dQuantOutras * dCustoOutras, "standard")
        
        If objEstoqueProduto.dQuantOutras3 = 0 Then
            dCustoOutras3 = 0
        End If
        QuantOutras3.Caption = Formata_Estoque(objEstoqueProduto.dQuantOutras3)
        CustoOutras3.Caption = Format(dCustoOutras3, FORMATO_CUSTO)
        ValorOutras3.Caption = Format(objEstoqueProduto.dQuantOutras3 * dCustoOutras3, "standard")
                
'        If giTipoVersao = VERSAO_FULL Then
            If objEstoqueProduto.dQuantConserto = 0 Then
                dCustoConserto = 0
            End If
            QuantConserto.Caption = Formata_Estoque(objEstoqueProduto.dQuantConserto)
            CustoConserto.Caption = Format(dCustoConserto, FORMATO_CUSTO)
            ValorConserto.Caption = Format(objEstoqueProduto.dQuantConserto * dCustoConserto, "standard")
            
            If objEstoqueProduto.dQuantConserto3 = 0 Then
                dCustoConserto3 = 0
            End If
            QuantConserto3.Caption = Formata_Estoque(objEstoqueProduto.dQuantConserto3)
            CustoConserto3.Caption = Format(dCustoConserto3, FORMATO_CUSTO)
            ValorConserto3.Caption = Format(objEstoqueProduto.dQuantConserto3 * dCustoConserto3, "standard")
            
            
            If objEstoqueProduto.dQuantDemo = 0 Then
                dCustoDemo = 0
            End If
            QuantDemo.Caption = Formata_Estoque(objEstoqueProduto.dQuantDemo)
            CustoDemo.Caption = Format(dCustoDemo, FORMATO_CUSTO)
            ValorDemo.Caption = Format(objEstoqueProduto.dQuantDemo * dCustoDemo, "standard")
            
            If objEstoqueProduto.dQuantDemo3 = 0 Then
                dCustoDemo3 = 0
            End If
            QuantDemo3.Caption = Formata_Estoque(objEstoqueProduto.dQuantDemo3)
            CustoDemo3.Caption = Format(dCustoDemo3, FORMATO_CUSTO)
            ValorDemo3.Caption = Format(objEstoqueProduto.dQuantDemo3 * dCustoDemo3, "standard")
            
            
            If objEstoqueProduto.dQuantBenef = 0 Then
                dCustoBenef = 0
            End If
            QuantBenef.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef)
            CustoBenef.Caption = Format(dCustoBenef, FORMATO_CUSTO)
            ValorBenef.Caption = Format(objEstoqueProduto.dQuantBenef * dCustoBenef, "standard")
            
            If objEstoqueProduto.dQuantBenef3 = 0 Then
                dCustoBenef3 = 0
            End If
            QuantBenef3.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef3)
            CustoBenef3.Caption = Format(dCustoBenef3, FORMATO_CUSTO)
            ValorBenef3.Caption = Format(objEstoqueProduto.dQuantBenef3 * dCustoBenef3, "standard")
 '       End If
        
        'Realiza a chamada a função que irá lê a Requisição e o Pedido em Compras
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQuantPedidaReqCompra, dQuantidadePedCompra, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 125323 And lErro <> 125328 Then gError 125331
        
        LabelQuantReqCompra = Formata_Estoque(dQuantPedidaReqCompra)
        LabelValReqCompra = Format(dQuantPedidaReqCompra * dCusto, "standard")
        LabelQuantPedCompra = Formata_Estoque(dQuantidadePedCompra)
        LabelValPedCompra = Format(dQuantidadePedCompra * dCusto, "standard")
        
    Else
    
        LocalizacaoFisica.Text = ""
        ContaContabil.PromptInclude = False
        ContaContabil.Text = ""
        ContaContabil.PromptInclude = True
        
        QuantDispTotal.Caption = ""
        QuantInd.Caption = ""
        QuantReservada.Caption = ""
        QuantDefeituosa.Caption = ""
        QuantTotal.Caption = ""
        QuantEmpenhada.Caption = ""
        QuantOP.Caption = ""
        QuantPedido.Caption = ""
        LabelQuantReqCompra.Caption = ""
        LabelQuantPedCompra.Caption = ""
        DataInventario.Caption = ""
        QuantConsig.Caption = ""
        QuantConsig3.Caption = ""
        ValorDispTotal.Caption = ""
        ValorInd.Caption = ""
        ValorReservada.Caption = ""
        ValorDefeituosa.Caption = ""
        ValorTotal.Caption = ""
        ValorEmpenhada.Caption = ""
        ValorOP.Caption = ""
        ValorPedido.Caption = ""
        ValorConsig.Caption = ""
        ValorConsig3.Caption = ""
        ValorOutras.Caption = ""
        ValorOutras3.Caption = ""
        LabelValReqCompra.Caption = ""
        LabelValPedCompra.Caption = ""
        CustoConsig.Caption = ""
        CustoConsig3.Caption = ""
        CustoOutras.Caption = ""
        CustoOutras3.Caption = ""
                
'        If giTipoVersao = VERSAO_FULL Then
        
            QuantConserto.Caption = ""
            QuantConserto.Caption = ""
            QuantConserto3.Caption = ""
            QuantDemo.Caption = ""
            QuantDemo3.Caption = ""
            QuantOutras.Caption = ""
            QuantOutras3.Caption = ""
            QuantBenef.Caption = ""
            QuantBenef3.Caption = ""
            ValorConserto.Caption = ""
            ValorConserto3.Caption = ""
            ValorDemo.Caption = ""
            ValorDemo3.Caption = ""
            ValorBenef.Caption = ""
            ValorBenef3.Caption = ""
            CustoConserto.Caption = ""
            CustoConserto3.Caption = ""
            CustoDemo.Caption = ""
            CustoDemo3.Caption = ""
            CustoBenef.Caption = ""
            CustoBenef3.Caption = ""
        
'        End If
        
    End If

    EstoqueProduto_TrazTela = SUCESSO

    Exit Function

Erro_EstoqueProduto_TrazTela:

    EstoqueProduto_TrazTela = gErr
    
    Select Case gErr

        Case 20791, 49943, 52980, 61123, 64015, 69837, 69838, 33630
        
        Case 49546
             lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
        
        Case 61124
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_INEXISTENTE", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case 125331
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155810)

    End Select

    Exit Function

End Function

Function CalculaValor_Empresa_Toda(sProduto As String, objEstoqueProdutoSUM As ClassEstoqueProduto, dCusto As Double, dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double)
'Calcula o Custo para a Empresa Toda

Dim lErro As Long
Dim objFilial As New AdmFiliais
Dim dQuantDispEmpresa As Double
Dim dValorDispEmpresa As Double
Dim dQuantConsigEmpresa As Double
Dim dValorConsigEmpresa As Double
Dim dQuantDemoEmpresa As Double
Dim dValorDemoEmpresa As Double
Dim dQuantBenefEmpresa As Double
Dim dValorBenefEmpresa As Double
Dim dQuantConsertoEmpresa As Double
Dim dValorConsertoEmpresa As Double
Dim dQuantOutrasEmpresa As Double
Dim dValorOutrasEmpresa As Double
Dim dQuantConsigEmpresa3 As Double
Dim dValorConsigEmpresa3 As Double
Dim dQuantDemoEmpresa3 As Double
Dim dValorDemoEmpresa3 As Double
Dim dQuantBenefEmpresa3 As Double
Dim dValorBenefEmpresa3 As Double
Dim dQuantConsertoEmpresa3 As Double
Dim dValorConsertoEmpresa3 As Double
Dim dQuantOutrasEmpresa3 As Double
Dim dValorOutrasEmpresa3 As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim iMesAnterior As Integer
Dim iAnoAnterior As Integer
Dim iCont As Integer
Dim iContAviso As Integer

On Error GoTo Erro_CalculaValor_Empresa_Toda
    
    'Varre Filial para Acumular as Quantidades e Valores = Quantidade * Custo
    For Each objFilial In gcolFiliais
        
        If objFilial.iCodFilial <> EMPRESA_TODA Then
            
            'Emite uma mensagem se as Filiais não estão apuradas no mesmo Mês-Ano.
            iCont = iCont + 1
            objEstoqueMes.iFilialEmpresa = objFilial.iCodFilial
                               
            'Le o último mes apurado
            lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
            If lErro <> SUCESSO And lErro <> 46225 Then gError 69881
            
            If iCont > 1 And iContAviso = 0 Then
                If iMesAnterior <> objEstoqueMes.iMes Or iAnoAnterior <> objEstoqueMes.iAno Then
                    Call Rotina_Aviso(vbOKOnly, "AVISO_FILIAIS_APURADAS_MESANO_DIFERENTES")
                    iContAviso = 1
                End If
            End If
                
            iMesAnterior = objEstoqueMes.iMes
            iAnoAnterior = objEstoqueMes.iAno
            
            objEstoqueProduto.sProduto = sProduto
            
            lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, objFilial.iCodFilial)
            If lErro <> SUCESSO And lErro <> 64014 Then gError 69839
                    
            lErro = Calcula_Valor_Tela(sProduto, objEstoqueProdutoSUM, dCusto, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, objFilial.iCodFilial)
            If lErro <> SUCESSO Then gError 69840
                    
            dQuantDispEmpresa = dQuantDispEmpresa + (objEstoqueProduto.dQuantDispNossa + objEstoqueProduto.dQuantInd + objEstoqueProduto.dQuantReservada + objEstoqueProduto.dQuantDefeituosa)
            dValorDispEmpresa = dValorDispEmpresa + (objEstoqueProduto.dQuantDispNossa + objEstoqueProduto.dQuantInd + objEstoqueProduto.dQuantReservada + objEstoqueProduto.dQuantDefeituosa) * dCusto
            
            dQuantConsigEmpresa = dQuantConsigEmpresa + objEstoqueProduto.dQuantConsig
            dValorConsigEmpresa = dValorConsigEmpresa + objEstoqueProduto.dQuantConsig * dCusto
            
            dQuantConsertoEmpresa = dQuantConsertoEmpresa + objEstoqueProduto.dQuantConserto
            dValorConsertoEmpresa = dValorConsertoEmpresa + objEstoqueProduto.dQuantConserto * dCusto
            
            dQuantBenefEmpresa = dQuantBenefEmpresa + objEstoqueProduto.dQuantBenef
            dValorBenefEmpresa = dValorBenefEmpresa + objEstoqueProduto.dQuantBenef * dCusto
            
            dQuantDemoEmpresa = dQuantDemoEmpresa + objEstoqueProduto.dQuantDemo
            dValorDemoEmpresa = dValorDemoEmpresa + objEstoqueProduto.dQuantDemo * dCusto
            
            dQuantOutrasEmpresa = dQuantOutrasEmpresa + objEstoqueProduto.dQuantOutras
            dValorOutrasEmpresa = dValorOutrasEmpresa + objEstoqueProduto.dQuantOutras * dCusto
            
            dQuantConsigEmpresa3 = dQuantConsigEmpresa3 + objEstoqueProduto.dQuantConsig3
            dValorConsigEmpresa3 = dValorConsigEmpresa3 + objEstoqueProduto.dQuantConsig3 * dCusto
            
            dQuantConsertoEmpresa3 = dQuantConsertoEmpresa3 + objEstoqueProduto.dQuantConserto3
            dValorConsertoEmpresa3 = dValorConsertoEmpresa3 + objEstoqueProduto.dQuantConserto3 * dCusto
            
            dQuantBenefEmpresa3 = dQuantBenefEmpresa3 + objEstoqueProduto.dQuantBenef3
            dValorBenefEmpresa3 = dValorBenefEmpresa3 + objEstoqueProduto.dQuantBenef3 * dCusto
            
            dQuantDemoEmpresa3 = dQuantDemoEmpresa3 + objEstoqueProduto.dQuantDemo3
            dValorDemoEmpresa3 = dValorDemoEmpresa3 + objEstoqueProduto.dQuantDemo3 * dCusto
            
            dQuantOutrasEmpresa3 = dQuantOutrasEmpresa3 + objEstoqueProduto.dQuantOutras3
            dValorOutrasEmpresa3 = dValorOutrasEmpresa3 + objEstoqueProduto.dQuantOutras3 * dCusto
                                            
        End If
        
    Next
        
    'Calcula o Custo que é o Valor Acumulado / Quantidade Acumulada
    If dQuantDispEmpresa > 0 Then
        dCusto = dValorDispEmpresa / dQuantDispEmpresa
    Else
        dCusto = 0
    End If
    
    
    If dQuantDemoEmpresa > 0 Then
        dCustoDemo = dValorDemoEmpresa / dQuantDemoEmpresa
    Else
        dCustoDemo = 0
    End If
        
    If dQuantBenefEmpresa > 0 Then
        dCustoBenef = dValorBenefEmpresa / dQuantBenefEmpresa
    Else
        dCustoBenef = 0
    End If
    
    If dQuantConsigEmpresa > 0 Then
        dCustoConsig = dValorConsigEmpresa / dQuantConsigEmpresa
    Else
        dCustoConsig = 0
    End If
        
    If dQuantConsertoEmpresa > 0 Then
        dCustoConserto = dValorConsertoEmpresa / dQuantConsertoEmpresa
    Else
        dCustoConserto = 0
    End If
    
    If dQuantOutrasEmpresa > 0 Then
        dCustoOutras = dValorOutrasEmpresa / dQuantOutrasEmpresa
    Else
        dCustoOutras = 0
    End If
    
    If dQuantDemoEmpresa3 > 0 Then
        dCustoDemo3 = dValorDemoEmpresa3 / dQuantDemoEmpresa3
    Else
        dCustoDemo3 = 0
    End If
    
    If dQuantBenefEmpresa3 > 0 Then
        dCustoBenef3 = dValorBenefEmpresa3 / dQuantBenefEmpresa3
    Else
        dCustoBenef3 = 0
    End If
    
    If dQuantConsigEmpresa3 > 0 Then
        dCustoConsig3 = dValorConsigEmpresa3 / dQuantConsigEmpresa3
    Else
        dCustoConsig3 = 0
    End If
    
    If dQuantConsertoEmpresa3 > 0 Then
        dCustoConserto3 = dValorConsertoEmpresa3 / dQuantConsertoEmpresa3
    Else
        dCustoConserto3 = 0
    End If
    
    If dQuantOutrasEmpresa3 > 0 Then
        dCustoOutras3 = dValorOutrasEmpresa3 / dQuantOutrasEmpresa3
    Else
        dCustoOutras3 = 0
    End If
    
    CalculaValor_Empresa_Toda = SUCESSO
    
    Exit Function
    
Erro_CalculaValor_Empresa_Toda:

    CalculaValor_Empresa_Toda = gErr
    
    Select Case gErr
    
        Case 69839, 69840, 69881
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155811)

    End Select

    Exit Function
    
End Function

Function Calcula_Valor_Tela(sProduto As String, objEstoqueProdutoSUM As ClassEstoqueProduto, dCusto As Double, dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double, iFilialEmpresa As Integer) As Long
'devolve o custo do produto e dos Escaninhos

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Calcula_Valor_Tela
    
    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 52122
    
    'O valor da diferenca vai depender se é um produto Produzido ou Compravel
    If objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
                
        'Le o Custo do Disponivel
        'Le o ultimo CustoMedio do Mes-Ano aberto
        lErro = CF("CustoMedioAtual_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 49944
        
        'Calcula o Custo para os Escaninhos
        If objEstoqueProdutoSUM.dQuantConserto <> 0 Then
            dCustoConserto = objEstoqueProdutoSUM.dValorConserto / objEstoqueProdutoSUM.dQuantConserto
        Else
            dCustoConserto = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConserto3 <> 0 Then
            dCustoConserto3 = objEstoqueProdutoSUM.dValorConserto3 / objEstoqueProdutoSUM.dQuantConserto3
        Else
            dCustoConserto3 = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConsig <> 0 Then
            dCustoConsig = objEstoqueProdutoSUM.dValorConsig / objEstoqueProdutoSUM.dQuantConsig
        Else
            dCustoConsig = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConsig3 <> 0 Then
            dCustoConsig3 = objEstoqueProdutoSUM.dValorConsig3 / objEstoqueProdutoSUM.dQuantConsig3
        Else
            dCustoConsig3 = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantDemo <> 0 Then
            dCustoDemo = objEstoqueProdutoSUM.dValorDemo / objEstoqueProdutoSUM.dQuantDemo
        Else
            dCustoDemo = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantDemo3 <> 0 Then
            dCustoDemo3 = objEstoqueProdutoSUM.dValorDemo3 / objEstoqueProdutoSUM.dQuantDemo3
        Else
            dCustoDemo3 = 0
        End If
        
        'Outras
        If objEstoqueProdutoSUM.dQuantOutras <> 0 Then
            dCustoOutras = objEstoqueProdutoSUM.dValorOutras / objEstoqueProdutoSUM.dQuantOutras
        Else
            dCustoOutras = 0
        End If

        If objEstoqueProdutoSUM.dQuantOutras3 <> 0 Then
            dCustoOutras3 = objEstoqueProdutoSUM.dValorOutras3 / objEstoqueProdutoSUM.dQuantOutras3
        Else
            dCustoOutras3 = 0
        End If
                
        'Beneficiamento
        If objEstoqueProdutoSUM.dQuantBenef <> 0 Then
            dCustoBenef = objEstoqueProdutoSUM.dValorBenef / objEstoqueProdutoSUM.dQuantBenef
        Else
            dCustoBenef = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantBenef3 <> 0 Then
            dCustoBenef3 = objEstoqueProdutoSUM.dValorBenef3 / objEstoqueProdutoSUM.dQuantBenef3
        Else
            dCustoBenef3 = 0
        End If
        
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
            
        'le o ultimo custo medio producao apurado
        lErro = CF("CustoMedioProducaoApurado_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69867 And lErro <> 69868 Then gError 49945
                            
        'Le os custos dos Escaninhos
        lErro = CF("CustoMedioProducaoApurado_Le_Escaninhos", objProduto.sCodigo, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 69790
    
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
    
        'le o custo standard corrente
        lErro = CF("CustoStandard_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69852 Then gError 55043
           
        'Le o Custo para os Escaninhos
        'Conserto
        If objEstoqueProdutoSUM.dQuantConserto <> 0 Then
            dCustoConserto = objEstoqueProdutoSUM.dValorConserto / objEstoqueProdutoSUM.dQuantConserto
        Else
            dCustoConserto = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConserto3 <> 0 Then
            dCustoConserto3 = objEstoqueProdutoSUM.dValorConserto3 / objEstoqueProdutoSUM.dQuantConserto3
        Else
            dCustoConserto3 = 0
        End If
            
        'Consignação
        If objEstoqueProdutoSUM.dQuantConsig <> 0 Then
            dCustoConsig = objEstoqueProdutoSUM.dValorConsig / objEstoqueProdutoSUM.dQuantConsig
        Else
            dCustoConsig = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConsig3 <> 0 Then
            dCustoConsig3 = objEstoqueProdutoSUM.dValorConsig3 / objEstoqueProdutoSUM.dQuantConsig3
        Else
            dCustoConsig3 = 0
        End If
        
        'Demostração
        If objEstoqueProdutoSUM.dQuantDemo <> 0 Then
            dCustoDemo = objEstoqueProdutoSUM.dValorDemo / objEstoqueProdutoSUM.dQuantDemo
        Else
            dCustoDemo = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantDemo3 <> 0 Then
            dCustoDemo3 = objEstoqueProdutoSUM.dValorDemo3 / objEstoqueProdutoSUM.dQuantDemo3
        Else
            dCustoDemo3 = 0
        End If
        
        'Outras
        If objEstoqueProdutoSUM.dQuantOutras <> 0 Then
            dCustoOutras = objEstoqueProdutoSUM.dValorOutras / objEstoqueProdutoSUM.dQuantOutras
        Else
            dCustoOutras = 0
        End If

        If objEstoqueProdutoSUM.dQuantOutras3 <> 0 Then
            dCustoOutras3 = objEstoqueProdutoSUM.dValorOutras3 / objEstoqueProdutoSUM.dQuantOutras3
        Else
            dCustoOutras3 = 0
        End If
                
        'Beneficiamento
        If objEstoqueProdutoSUM.dQuantBenef <> 0 Then
            dCustoBenef = objEstoqueProdutoSUM.dValorBenef / objEstoqueProdutoSUM.dQuantBenef
        Else
            dCustoBenef = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantBenef3 <> 0 Then
            dCustoBenef3 = objEstoqueProdutoSUM.dValorBenef3 / objEstoqueProdutoSUM.dQuantBenef3
        Else
            dCustoBenef3 = 0
        End If
    
    End If
    
    Calcula_Valor_Tela = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valor_Tela:
            
    Calcula_Valor_Tela = gErr
    
    Select Case gErr
    
        Case 49944, 49945, 52122, 55043, 69790
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155812)

    End Select

    Exit Function
    
End Function

Private Sub Limpa_Campos_EstoqueProduto1()

Dim iIndice As Integer

    For iIndice = 0 To 1
        AlmoxarifadoLabel(iIndice).Caption = ""
    Next

    For iIndice = 0 To 1
            
        UnidMedLabel(iIndice).Caption = ""
        ProdutoLabel(iIndice).Caption = ""
    Next

    DataInventario.Caption = ""

    Almoxarifado.Clear
    Padrao.Value = DESMARCADO
    LocalizacaoFisica.Text = ""
    
    ContaContabil.PromptInclude = False
    ContaContabil.Text = ""
    ContaContabil.PromptInclude = True
    QuantDispTotal.Caption = ""
    QuantInd.Caption = ""
    QuantReservada.Caption = ""
    QuantDefeituosa.Caption = ""
    QuantTotal.Caption = ""
    QuantEmpenhada.Caption = ""
    QuantOP.Caption = ""
    QuantPedido.Caption = ""
    LabelQuantReqCompra.Caption = ""
    LabelQuantPedCompra.Caption = ""
    DataInventario.Caption = ""
    QuantConsig.Caption = ""
    QuantConsig3.Caption = ""
    QuantOutras.Caption = ""
    QuantOutras3.Caption = ""
    ValorDispTotal.Caption = ""
    ValorInd.Caption = ""
    ValorReservada.Caption = ""
    ValorDefeituosa.Caption = ""
    ValorTotal.Caption = ""
    ValorEmpenhada.Caption = ""
    ValorOP.Caption = ""
    ValorPedido.Caption = ""
    ValorConsig.Caption = ""
    ValorConsig3.Caption = ""
    ValorOutras.Caption = ""
    ValorOutras3.Caption = ""
    LabelValReqCompra.Caption = ""
    LabelValPedCompra.Caption = ""
    
'    If giTipoVersao = VERSAO_FULL Then
        QuantConserto.Caption = ""
        QuantConserto3.Caption = ""
        QuantDemo.Caption = ""
        QuantDemo3.Caption = ""
        QuantBenef.Caption = ""
        QuantBenef3.Caption = ""
        ValorConserto.Caption = ""
        ValorConserto3.Caption = ""
        ValorDemo.Caption = ""
        ValorDemo3.Caption = ""
        ValorBenef.Caption = ""
        ValorBenef3.Caption = ""
'    End If
    
End Sub

'Private Function Produto_Le_Quantidades_Compras(objProduto As ClassProduto, dQuantPedidaReqCompra As Double, dQuantidadePedCompra As Double, ByVal iFilialEmpresa As Integer) As Long
''Realiza a leitura da QuantidadePedida na Requisição e a Quantidade no Pedido de Compras
'
'Dim lErro As Long
'Dim alComando(1 To 2) As Long
'Dim iIndice As Integer
'Dim dQuantidade As Double
'Dim dFator As Double
'Dim sUM As String
'
'On Error GoTo Erro_Produto_Le_Quantidades_Compras
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir
'        If alComando(iIndice) = 0 Then gError 125320
'    Next
'
'    dQuantPedidaReqCompra = 0
'    dQuantidadePedCompra = 0
'
'    'Inicializa as strings
'    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'    'Busca no BD o código passado
'    lErro = Comando_Executar(alComando(1), "SELECT ItensReqCompra.UM, SUM(ItensReqCompra.Quantidade - ItensReqCompra.QuantPedida - ItensReqCompra.QuantRecebida - ItensReqCompra.QuantCancelada) FROM Produtos, ProdutosFilial, ItensReqCompra, RequisicaoCompra WHERE Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.Produto = ItensReqCompra.Produto AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.DataEnvio <> ? AND Produtos.Codigo = ? AND ProdutosFilial.FilialEmpresa = ? GROUP BY ItensReqCompra.Produto, ItensReqCompra.UM", sUM, dQuantidade, DATA_NULA, objProduto.sCodigo, iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 125321
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125322
'
'    'Se nao encontrou ==> erro
''    If lErro <> AD_SQL_SUCESSO Then gError 125323
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Realiza a converção para um só UM
'        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 125324
'
'        'Preenche a variável quantidade
'        dQuantPedidaReqCompra = dQuantPedidaReqCompra + (dQuantidade * dFator)
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125325
'
'    Loop
'
'    'Inicializa as strings
'    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'    'Busca no BD o pedido com o produto passado
'    lErro = Comando_Executar(alComando(2), "SELECT ItensPedCompra.UM, SUM(ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento) FROM Produtos, ProdutosFilial, ItensPedCompra, PedidoCompra WHERE Produtos.Codigo = ProdutosFilial.Produto And ProdutosFilial.Produto = ItensPedCompra.Produto AND ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND PedidoCompra.DataEnvio <> ? AND Produtos.Codigo = ? AND ProdutosFilial.FilialEmpresa = ? GROUP BY ItensPedCompra.Produto, ItensPedCompra.UM", sUM, dQuantidade, DATA_NULA, objProduto.sCodigo, iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 125326
'
'    lErro = Comando_BuscarPrimeiro(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125327
'
'    'Se nao encontrou ==> erro
''    If lErro <> AD_SQL_SUCESSO Then gError 125328
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Realiza a converção para uma mesma UM
'        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 125329
'
'        'Preenche a quantidade
'        dQuantidadePedCompra = dQuantidadePedCompra + (dQuantidade * dFator)
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125330
'
'    Loop
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Produto_Le_Quantidades_Compras = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Le_Quantidades_Compras:
'
'    Produto_Le_Quantidades_Compras = gErr
'
'    Select Case gErr
'
'        Case 125320
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
'
'        Case 125321, 125322, 125325, 125326, 125327, 125330
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objProduto.sCodigo)
'
'        Case 125323, 125328 'Não encotrou o elemento
'
'        Case 125324, 125329
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155813)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PRODUTOS_ALMOXARIFADO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Produto X Almoxarifado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "EstoqueProduto"
    
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
        If Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel1_Click
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


