VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTEstoqueInicial"
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

Const TAB_PRINCIPAL = 1
Const TAB_COMPLEMENTAR = 2

Dim iFrameAtual As Integer
Public iAlterado As Integer
Public iAlmoxarifadoTrocou As Integer
Public gcolRastreamento As Collection
Private gsProduto As String

Private gobjGenerico As AdmGenerico

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifado As AdmEvento
Attribute objEventoAlmoxarifado.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProduto As AdmEvento
Attribute objEventoEstoqueProduto.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get ListaAlmoxarifado() As Object
     Set ListaAlmoxarifado = objUserControl.Controls("ListaAlmoxarifado")
End Property

Public Property Get BotaoEstoquesIniciais() As Object
     Set BotaoEstoquesIniciais = objUserControl.Controls("BotaoEstoquesIniciais")
End Property

Public Property Get LocalizacaoFisica() As Object
     Set LocalizacaoFisica = objUserControl.Controls("LocalizacaoFisica")
End Property

Public Property Get Padrao() As Object
     Set Padrao = objUserControl.Controls("Padrao")
End Property

Public Property Get Fixar() As Object
     Set Fixar = objUserControl.Controls("Fixar")
End Property

Public Property Get UpDown1() As Object
     Set UpDown1 = objUserControl.Controls("UpDown1")
End Property

Public Property Get TvwContas() As Object
     Set TvwContas = objUserControl.Controls("TvwContas")
End Property

Public Property Get DataInicial() As Object
     Set DataInicial = objUserControl.Controls("DataInicial")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

'Public Property Get TvwProdutos() As Object
'     Set TvwProdutos = objUserControl.Controls("TvwProdutos")
'End Property

Public Property Get UnidMed() As Object
     Set UnidMed = objUserControl.Controls("UnidMed")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get AlmoxarifadoLabel() As Object
     Set AlmoxarifadoLabel = objUserControl.Controls("AlmoxarifadoLabel")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get LabelAlmoxarifado() As Object
     Set LabelAlmoxarifado = objUserControl.Controls("LabelAlmoxarifado")
End Property

Public Property Get LabelContas() As Object
     Set LabelContas = objUserControl.Controls("LabelContas")
End Property

'Public Property Get LabelProduto() As Object
'     Set LabelProduto = objUserControl.Controls("LabelProduto")
'End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get QuantConserto() As Object
     Set QuantConserto = objUserControl.Controls("QuantConserto")
End Property

Public Property Get ValorConserto() As Object
     Set ValorConserto = objUserControl.Controls("ValorConserto")
End Property

Public Property Get QuantConsig() As Object
     Set QuantConsig = objUserControl.Controls("QuantConsig")
End Property

Public Property Get ValorConsig() As Object
     Set ValorConsig = objUserControl.Controls("ValorConsig")
End Property

Public Property Get QuantDemo() As Object
     Set QuantDemo = objUserControl.Controls("QuantDemo")
End Property

Public Property Get ValorDemo() As Object
     Set ValorDemo = objUserControl.Controls("ValorDemo")
End Property

Public Property Get QuantOutras() As Object
     Set QuantOutras = objUserControl.Controls("QuantOutras")
End Property

Public Property Get ValorOutras() As Object
     Set ValorOutras = objUserControl.Controls("ValorOutras")
End Property

Public Property Get QuantBenef() As Object
     Set QuantBenef = objUserControl.Controls("QuantBenef")
End Property

Public Property Get ValorBenef() As Object
     Set ValorBenef = objUserControl.Controls("ValorBenef")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get QuantConserto3() As Object
     Set QuantConserto3 = objUserControl.Controls("QuantConserto3")
End Property

Public Property Get ValorConserto3() As Object
     Set ValorConserto3 = objUserControl.Controls("ValorConserto3")
End Property

Public Property Get QuantConsig3() As Object
     Set QuantConsig3 = objUserControl.Controls("QuantConsig3")
End Property

Public Property Get ValorConsig3() As Object
     Set ValorConsig3 = objUserControl.Controls("ValorConsig3")
End Property

Public Property Get QuantDemo3() As Object
     Set QuantDemo3 = objUserControl.Controls("QuantDemo3")
End Property

Public Property Get ValorDemo3() As Object
     Set ValorDemo3 = objUserControl.Controls("ValorDemo3")
End Property

Public Property Get QuantOutras3() As Object
     Set QuantOutras3 = objUserControl.Controls("QuantOutras3")
End Property

Public Property Get ValorOutras3() As Object
     Set ValorOutras3 = objUserControl.Controls("ValorOutras3")
End Property

Public Property Get QuantBenef3() As Object
     Set QuantBenef3 = objUserControl.Controls("QuantBenef3")
End Property

Public Property Get ValorBenef3() As Object
     Set ValorBenef3 = objUserControl.Controls("ValorBenef3")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get AlmoxarifadoComplementar() As Object
     Set AlmoxarifadoComplementar = objUserControl.Controls("AlmoxarifadoComplementar")
End Property

Public Property Get ProdutoComplementar() As Object
     Set ProdutoComplementar = objUserControl.Controls("ProdutoComplementar")
End Property

Public Property Get UMComplementar() As Object
     Set UMComplementar = objUserControl.Controls("UMComplementar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoImportar() As Object
     Set BotaoImportar = objUserControl.Controls("BotaoImportar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoConsultar() As Object
     Set BotaoConsultar = objUserControl.Controls("BotaoConsultar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get BotaoRastro() As Object
     Set BotaoRastro = objUserControl.Controls("BotaoRastro")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Almoxarifado_GotFocus()

    LabelAlmoxarifado.Visible = True
    ListaAlmoxarifado.Visible = True
    TvwContas.Visible = False
    LabelContas.Visible = False
'    TvwProdutos.Visible = False
'    LabelProduto.Visible = False

End Sub

Public Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sAlmoxarifadoAntigo As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaEnxuta As String

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Guarda o conteúdo em Almoxarifado.text
        sAlmoxarifadoAntigo = Almoxarifado.Text
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 44922
        
        'se o almoxarifado não pertencer a filial em questão ==> erro
        If objAlmoxarifado.iFilialEmpresa <> giFilialEmpresa Then Error 55868
        
        'Se o Almoxarifado foi Alterado
        If sAlmoxarifadoAntigo <> Almoxarifado.Text Then
            
            'preenche a conta contabil
            If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
                        
                If Len(Trim(objAlmoxarifado.sContaContabil)) > 0 Then
                    sContaEnxuta = String(STRING_CONTA, 0)
                    
                    lErro = Mascara_RetornaContaEnxuta(objAlmoxarifado.sContaContabil, sContaEnxuta)
                    If lErro <> SUCESSO Then Error 61371
                    
                    ContaContabil.PromptInclude = False
                    ContaContabil.Text = sContaEnxuta
                    ContaContabil.PromptInclude = True
                Else
                    ContaContabil.PromptInclude = False
                    ContaContabil.Text = ""
                    ContaContabil.PromptInclude = True
                End If
                
           End If
        
        End If
    
    End If
    
    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True


    Select Case Err

        Case 44922

        Case 55868
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIAL_DIFERENTE", Err, objAlmoxarifado.iCodigo, giFilialEmpresa)
        
        Case 61370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objAlmoxarifado.sContaContabil)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 155732)

    End Select

    Exit Sub

End Sub

Public Sub AlmoxarifadoLabel_Click()

Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colSelecao As New Collection

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxarifado)

End Sub

Public Sub BotaoConsultar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoConsultar_Click

    'Verifica preenchimento de Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then Error 44930

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 44931

    lErro = Trata_Troca_Produto_Almoxarifado()
    If lErro <> SUCESSO Then Error 44932

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoConsultar_Click:

    Select Case Err
    
        Case 44930
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", Err)

        Case 44931
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case 44932
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155733)
        
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoEstoquesIniciais_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoEstoquesIniciais_Click

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 45916
        
        objEstoqueProduto.sProduto = sProdutoFormatado

    End If

    'Verifica se Almoxarifado está preenchido
    If Almoxarifado.Text <> "" Then
    
        'preenche o objAlmoxarifado
        objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO Then Error 45917

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    End If
    
    'Chama a tela EstoqueInicialLista
    Call Chama_Tela("EstoqueInicialLista", colSelecao, objEstoqueProduto, objEventoEstoqueProduto)

    Exit Sub
    
Erro_BotaoEstoquesIniciais_Click:

    Select Case Err
    
        Case 45916, 45917

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155734)
        
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoRastro_Click()

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoRastro_Click

    'Verifica preenchimento de Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then gError 71824

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 71825

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 71826
    
    If lErro = 25041 Then gError 71827
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 71828

    If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then gError 71829
    
    'Chama a tela de Rasteamento dos Produtos criando uma nova instancia da tela independente se já houver alguma associada a outra tela
    Call Chama_Tela_Nova_Instancia("RastroEstoqueInicial", Me, gobjGenerico)
        
    'Desabilita a tela de nota fiscal
    gobjGenerico.vVariavel = DESABILITA_TELA
        
    Exit Sub
    
Erro_BotaoRastro_Click:

    Select Case gErr
    
        Case 71824
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)

        Case 71825, 71828
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 71826
        
        Case 71827
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
        
        Case 71829
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155735)
        
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) = 0 Then Error 21310
    
    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then Error 21311

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 21312
    
    If lErro <> SUCESSO Then Error 28767
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27197

    'Preenche objEstoqueProduto
    objEstoqueProduto.sProduto = objProduto.sCodigo
    
    lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
    If lErro <> SUCESSO Then Error 44927
    
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    objEstoqueProduto.dtDataInicial = MaskedParaDate(DataInicial)

    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then Error 31220
    
    If lErro <> SUCESSO Then Error 28766

    'Pede confirmação para exclusão de Estoque Inicial
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ESTOQUEINICIAL", objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
    
    'Chama EstoqueInicial_Exclui
    lErro = CF("EstoqueInicial_Exclui", objEstoqueProduto)
    If lErro <> SUCESSO Then Error 21314

    Call Limpa_Tela_EstoqueInicial

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21310
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", Err)

        Case 21311, 27197
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 21312, 21314, 31220, 44927
            
        Case 28766
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo, objEstoqueProduto.iAlmoxarifado)
        
        Case 28767
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155736)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21315

    Call Limpa_Tela_EstoqueInicial

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21315

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155737)

    End Select

    Exit Sub

End Sub

Public Sub BotaoImportar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImportar_Click
        
    GL_objMDIForm.MousePointer = vbHourglass

    lErro = CF("ImportarEstoqueInicial")
    If lErro <> SUCESSO Then gError 129855

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImportar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 129855

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155738)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21316

    'Limpa a tela
    Call Limpa_Tela_EstoqueInicial

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21316

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155739)

    End Select

    Exit Sub

End Sub

Public Sub ContaContabil_GotFocus()
    
    LabelAlmoxarifado.Visible = False
    ListaAlmoxarifado.Visible = False
    TvwContas.Visible = True
    LabelContas.Visible = True
'    TvwProdutos.Visible = False
'    LabelProduto.Visible = False
    
End Sub

Public Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial, iAlterado)

End Sub

Public Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_DataInicial_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataInicial.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then Error 21318

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 21318

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155740)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Carrega índices da tela
    Call TelaIndice_Preenche(Me)

    'Retorno da tela de RastroProdNF
    If gobjGenerico.vVariavel = HABILITA_TELA Then
    
        'Habilita a tela de nota fiscal
        Me.Enabled = True
         
    Else
    
        'Desabilita a tela de nota fiscal
        Me.Enabled = False
    
    End If


End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    Set objEventoContaContabil = New AdmEvento
    Set objEventoAlmoxarifado = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoqueProduto = New AdmEvento

    Set gobjGenerico = New AdmGenerico

    gobjGenerico.vVariavel = HABILITA_TELA

    Set gcolRastreamento = New Collection

    iFrameAtual = 1
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 21319

    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta_Modulo", TvwContas.Nodes, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then Error 19161
    
    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 22160

    ContaContabil.Mask = sMascaraConta

'    'Inicializa a Lista de Produtos
'    lErro = CF("Carga_Arvore_Produto_Estoque",TvwProdutos.Nodes)
'    If lErro <> SUCESSO And lErro <> 58762 Then Error 21321
'    If lErro <> SUCESSO Then Error 59222

    'Preenche a ListBox Almoxarifado com os Almoxarifados da FilialEmpresa
    lErro = Carga_ListaAlmoxarifado()
    If lErro <> SUCESSO Then Error 44921
    
    DataInicial.Text = Format(gdtDataAtual, "dd/mm/yy")
        
    Quantidade.Format = FORMATO_ESTOQUE
    QuantConsig.Format = FORMATO_ESTOQUE
    QuantOutras.Format = FORMATO_ESTOQUE
    QuantConsig3.Format = FORMATO_ESTOQUE
    QuantOutras3.Format = FORMATO_ESTOQUE
    
'    If giTipoVersao = VERSAO_FULL Then
    
        QuantDemo.Format = FORMATO_ESTOQUE
        QuantConserto.Format = FORMATO_ESTOQUE
        QuantBenef.Format = FORMATO_ESTOQUE
        QuantConserto3.Format = FORMATO_ESTOQUE
        QuantDemo3.Format = FORMATO_ESTOQUE
        QuantBenef3.Format = FORMATO_ESTOQUE
    
'    End If
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 19161, 21319, 21321, 22160, 28444, 44921

        Case 59222
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_PRODUTOS_CTL_ESTOQUE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155741)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoContaContabil = Nothing
    
    Set objEventoAlmoxarifado = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoqueProduto = Nothing
   
    Set gobjGenerico = Nothing
    Set gcolRastreamento = Nothing
   
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
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
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 49541

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 49542

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 49543

        'conta não cadastrada
        If lErro = 5700 Then Error 49544

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case Err

    Case 49541, 49543

    Case 49542
        lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
        
    Case 49544
        lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155742)

    End Select

    Exit Sub

End Sub

Private Sub Label10_Click()

End Sub

Public Sub LocalizacaoFisica_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim sConta As String
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 49540

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 49540

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155743)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxarifado_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche AlmoxPadrao com o nome reduzido de objAlmoxarifado
    Almoxarifado.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

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
        If lErro <> SUCESSO Then Error 49537

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 49537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155744)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoqueProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sContaMascarada As String

On Error GoTo Erro_objEventoEstoqueProduto_evSelecao

    Set objEstoqueProduto = obj1
    
    lErro = Mascara_RetornaProdutoEnxuto(objEstoqueProduto.sProduto, sProduto)
    If lErro <> SUCESSO Then Error 45919

    'Coloca na tela o Produto selecionado
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
    
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    lErro = Trata_Troca_Produto_Almoxarifado()
    If lErro <> SUCESSO Then Error 44986

    iAlterado = 0

    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoEstoqueProduto_evSelecao:

    Select Case Err

        Case 44986

        Case 45919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155745)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then Error 45908

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 45909

    'Coloca na tela o Produto selecionado
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
    DescricaoProduto.Caption = objProduto.sDescricao
    UnidMed.Caption = objProduto.sSiglaUMEstoque
    Call CF2(Me, "EstIni_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)

    Call Produto_Validate(bSGECancelDummy)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 45910

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 45908, 45910

        Case 45909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155746)

    End Select

    Exit Sub

End Sub

Public Sub Padrao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Produto_GotFocus()
    
'    LabelAlmoxarifado.Visible = False
'    ListaAlmoxarifado.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    TvwProdutos.Visible = True
'    LabelProduto.Visible = True

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Produto_Validate

    If Produto.Text <> gsProduto Then
        
        Set gcolRastreamento = New Collection
        
    End If

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199355

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 33637

        'O Produto é Gerencial
        If lErro = 25043 Then gError 33638

        'O Produto não está cadastrado
        If lErro = 25041 Then gError 33639

        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 44935

        'Preenche ProdutoDescricao com Descrição do Produto
        DescricaoProduto.Caption = objProduto.sDescricao

        'Preenche a Unidade de Medida
        UnidMed.Caption = objProduto.sSiglaUMEstoque
        
        Call CF2(Me, "EstIni_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)
        
    End If
        
    gsProduto = Produto.Text
        
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 33637, 33638, 199355

        Case 33639
            'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                DescricaoProduto.Caption = ""
                UnidMed.Caption = ""
                Call CF2(Me, "EstIni_Produto_ExibeInfo", "")
                'Segura o foco
            End If

        Case 44935
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155747)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoLabel1_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLabel1_Click

    sSelecao = "ControleEstoque <> ? AND Gerencial = ? AND Ativo = ?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    colSelecao.Add NAO_GERENCIAL
    colSelecao.Add PRODUTO_ATIVO
    
    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83059
        
        objProduto.sCodigo = sProdutoFormatado

    End If
    
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_ProdutoLabel1_Click:

    Select Case gErr

        Case 83059

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155748)

    End Select

    Exit Sub


End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_Quantidade_Validate

    If Len(Trim(Quantidade.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 33636
        
    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, Quantidade.Text, ESCANINHO_DISPONIVEL, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 83006

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 83007

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 83006

        Case 83007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(Quantidade.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155749)

    End Select

    Exit Sub

End Sub

Public Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'Move os dados do Plano de Contas do banco de dados para a árvore colNodes.
        lErro = CF("Carga_Arvore_Conta_Modulo1", objNode, TvwContas.Nodes, MODULO_ESTOQUE)
        If lErro <> SUCESSO Then Error 43609
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 43609
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155750)
        
    End Select
        
    Exit Sub

End Sub

Public Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim sConta As String
Dim sContaEnxuta As String
Dim sContaFormatada As String
Dim sCaracterInicial As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_TvwContas_NodeClick
    
    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial = "A" Then

        sConta = right(Node.Key, Len(Node.Key) - 1)
        
        sContaEnxuta = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 43610
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    
        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 43611
                
        'Conta não cadastrada
        If lErro = 5700 Then Error 43612
    
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 43610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case 43611
    
        Case 43612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaContabil.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155751)
            
    End Select
        
    Exit Sub

End Sub

'Public Sub TvwProdutos_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos_Estoque",objNode, TvwProdutos.Nodes)
'        If lErro <> SUCESSO Then Error 48082
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case Err
'
'        Case 48082
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155752)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TvwProdutos_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sCodigo As String
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_Produtos_NodeClick
'
'    'Armazena key do nó clicado sem caracter inicial
'    sCodigo = Right(Node.Key, Len(Node.Key) - 1)
'
'    'Verifica se produto tem filhos
'    If Node.Children > 0 Then Exit Sub
'
'    objProduto.sCodigo = sCodigo
'
'    'Lê Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then Error 21324
'
'    'Verifica se Produto é gerencial
'    If lErro = SUCESSO And objProduto.iGerencial = GERENCIAL Then Exit Sub
'
'    lErro = Traz_Produto_Tela(sCodigo)
'    If lErro <> SUCESSO Then Error 21325
'
'    Exit Sub
'
'Erro_Produtos_NodeClick:
'
'    Select Case Err
'
'        Case 21324, 21325
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155753)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub Limpa_Tela_EstoqueInicial()

Dim iAlmoxarifado As Integer
Dim sProduto As String
Dim sData As String

    sProduto = Produto.Text

    If Fixar.Value = 1 Then

        'Função genérica que limpa campos da tela
        Call Limpa_Tela(Me)

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    Else

        'Função genérica que limpa campos da tela
        Call Limpa_Tela(Me)

        Padrao.Value = 0
        UnidMed.Caption = ""
        DescricaoProduto.Caption = ""
        Call CF2(Me, "EstIni_Produto_ExibeInfo", "")
        UMComplementar.Caption = ""
        ProdutoComplementar.Caption = ""
        AlmoxarifadoComplementar.Caption = ""
        
    End If

    DataInicial.Text = Format(gdtDataAtual, "dd/mm/yy")
    Set gcolRastreamento = New Collection

    iAlterado = 0

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_Tela_Extrai

    'tulio300102
    'alteracao para filtrar por filial
    'Informa tabela associada à Tela
    sTabela = "EstoqueProduto_Almoxarifado"

    'Lê os dados da Tela EstoqueProduto
    lErro = Move_Tela_Memoria(objEstoqueProduto)
    If lErro <> SUCESSO Then Error 21338

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objEstoqueProduto.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Almoxarifado", objEstoqueProduto.iAlmoxarifado, 0, "Almoxarifado"

    'tulio300102
    'alteracao para filtrar por filial
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err
    
        Case 21338

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155754)

    End Select

    Exit Sub

End Sub

Public Function Traz_Produto_Tela(sCodigo As String) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Produto_Tela

    lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then Error 34677

    Call Produto_Validate(bSGECancelDummy)
    
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err

        Case 34677

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155755)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodigo As String
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Tela_Preenche

    objEstoqueProduto.sProduto = colCampoValor.Item("Produto").vValor
    objEstoqueProduto.iAlmoxarifado = colCampoValor.Item("Almoxarifado").vValor

    If (objEstoqueProduto.sProduto <> "") And (objEstoqueProduto.iAlmoxarifado <> 0) Then
        
        objProduto.sCodigo = objEstoqueProduto.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 42397
        
        If lErro <> SUCESSO Then Error 35996

        lErro = EstoqueInicial_TrazTela(objProduto, objEstoqueProduto.iAlmoxarifado, gcolRastreamento)
        If lErro <> SUCESSO Then Error 44925
        
        'Preenche objProdutoFilial
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        objProdutoFilial.sProduto = objProduto.sCodigo
        
        'Lê ProdutoFilial
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then Error 55203
    
        If lErro = SUCESSO Then
            If objProdutoFilial.iAlmoxarifado = objEstoqueProduto.iAlmoxarifado Then
                Padrao.Value = MARCADO
            Else
                Padrao.Value = DESMARCADO
            End If
        End If
        
    End If
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 35996, 42397, 44925, 55203

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155756)

    End Select

    Exit Sub

End Sub

Private Function EstoqueInicial_TrazTela(objProduto As ClassProduto, iAlmoxarifado As Integer, colRastreamento As Collection) As Long
'Mostra os dados do estoque do produto na tela

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iIndice As Integer
Dim bCancel As Boolean
Dim sContaMascarada As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_EstoqueInicial_TrazTela

    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado

    'Lê o Estoque Produto
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 21326

    'se o estoqueproduto não está cadastrado ==> erro
    If lErro = 21306 Then gError 44934
    
    objAlmoxarifado.iCodigo = objEstoqueProduto.iAlmoxarifado
    
    'Lê o Almoxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 44923

    'Se não achou o Almoxarifado --> erro
    If lErro = 25056 Then gError 44926

    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True

    UnidMed.Caption = objProduto.sSiglaUMEstoque
    
    DescricaoProduto.Caption = objProduto.sDescricao
    Call CF2(Me, "EstIni_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)

    Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    LocalizacaoFisica.Text = objEstoqueProduto.sLocalizacaoFisica
        
    AlmoxarifadoComplementar.Caption = Almoxarifado.Text
    ProdutoComplementar.Caption = Produto.Text
    UMComplementar = UnidMed.Caption
    
    'preenche a conta contabil
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        If objEstoqueProduto.sContaContabil <> "" Then
            sContaMascarada = String(STRING_CONTA, 0)
    
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaMascarada)
            If lErro <> SUCESSO Then gError 49545
        Else
            sContaMascarada = ""
        End If
    
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True
    
        Call ContaContabil_Validate(bCancel)
    
    End If
    
    'preenche quantidade inicial
    If objEstoqueProduto.dQuantidadeInicial <> 0 Then
        Quantidade.Text = objEstoqueProduto.dQuantidadeInicial
    Else
        Quantidade.Text = 0
    End If
    
    If objEstoqueProduto.dSaldoInicial <> 0 Then
        Valor.Text = CStr(objEstoqueProduto.dSaldoInicial)
    Else
        Valor.Text = "0"
    End If

    DataInicial.Text = Format(objEstoqueProduto.dtDataInicial, "dd/mm/yy")

    'Consig
    QuantConsig.Text = objEstoqueProduto.dQuantInicialConsig
    ValorConsig.Text = CStr(objEstoqueProduto.dValorInicialConsig)

    'Outras
    QuantOutras.Text = objEstoqueProduto.dQuantInicialOutras
    ValorOutras.Text = CStr(objEstoqueProduto.dValorInicialOutras)

    'Consig3
    QuantConsig3.Text = objEstoqueProduto.dQuantInicialConsig3
    ValorConsig3.Text = CStr(objEstoqueProduto.dValorInicialConsig3)

    'Outras3
    QuantOutras3.Text = objEstoqueProduto.dQuantInicialOutras3
    ValorOutras3.Text = CStr(objEstoqueProduto.dValorInicialOutras3)

'    If giTipoVersao = VERSAO_FULL Then
    
        'Conserto
        QuantConserto.Text = objEstoqueProduto.dQuantInicialConserto
        ValorConserto.Text = CStr(objEstoqueProduto.dValorInicialConserto)
    
        'Demo
        QuantDemo.Text = objEstoqueProduto.dQuantInicialDemo
        ValorDemo.Text = CStr(objEstoqueProduto.dValorInicialDemo)
    
        'Benef
        QuantBenef.Text = objEstoqueProduto.dQuantInicialBenef
        ValorBenef.Text = CStr(objEstoqueProduto.dValorInicialBenef)
    
        'Conserto3
        QuantConserto3.Text = objEstoqueProduto.dQuantInicialConserto3
        ValorConserto3.Text = CStr(objEstoqueProduto.dValorInicialConserto3)


        'Demo3
        QuantDemo3.Text = objEstoqueProduto.dQuantInicialDemo3
        ValorDemo3.Text = CStr(objEstoqueProduto.dValorInicialDemo3)
    
        'Benef3
        QuantBenef3.Text = objEstoqueProduto.dQuantInicialBenef3
        ValorBenef3.Text = CStr(objEstoqueProduto.dValorInicialBenef3)
    
'    End If
    
    Set colRastreamento = New Collection
    
    lErro = CF("RastroEstIni_Le", objProduto.sCodigo, iAlmoxarifado, colRastreamento)
    If lErro <> SUCESSO Then gError 71830
    
    EstoqueInicial_TrazTela = SUCESSO

    Exit Function

Erro_EstoqueInicial_TrazTela:

    EstoqueInicial_TrazTela = gErr

    Select Case gErr

        Case 21326, 44923, 71830
        
        Case 44926
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case 44934
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case 49545
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155757)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iAlmoxarifadoPadrao As Integer
Dim vbMsgRes As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then Error 21297

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 21298

    'Verifica preenchimento de Data
    If Len(Trim(DataInicial.ClipText)) = 0 Then Error 28808
    
    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objEstoqueProduto)
    If lErro <> SUCESSO Then Error 16932
        
    lErro = Trata_Alteracao(objEstoqueProduto, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
    If lErro <> SUCESSO Then Error 32327
    
    'verifica se o produto é rastreado e se for, verifica se todos os totais estão rastreados
    lErro = Testa_Rastreamento(objEstoqueProduto, gcolRastreamento)
    If lErro <> SUCESSO Then Error 20812
    
    'Almoxarifado padrão
    If Padrao.Value <> NAO_PADRAO Then
        iAlmoxarifadoPadrao = MARCADO
    Else
        iAlmoxarifadoPadrao = DESMARCADO
    End If
    
    'Chama EstoqueInicial_Grava
    lErro = CF("EstoqueInicial_Grava", objEstoqueProduto, iAlmoxarifadoPadrao, gcolRastreamento)
    If lErro <> SUCESSO Then Error 28807


    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 16932, 20812, 32327

        Case 21297
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", Err)

        Case 21298
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 28807

        Case 28808
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155758)

    End Select

    Exit Function

End Function

Private Function Testa_Rastreamento(objEstoqueProduto As ClassEstoqueProduto, colRastreamento As Collection) As Long

Dim lErro As Long
Dim objRastroEstIni As ClassRastroEstIni
Dim dQuantDisp As Double
Dim dQuantConserto As Double
Dim dQuantConsig As Double
Dim dQuantDemo As Double
Dim dQuantOutros As Double
Dim dQuantBenef As Double
Dim dQuantConserto3 As Double
Dim dQuantConsig3 As Double
Dim dQuantDemo3 As Double
Dim dQuantOutros3 As Double
Dim dQuantBenef3 As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Testa_Rastreamento

    objProduto.sCodigo = objEstoqueProduto.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 100001
    
    If lErro <> SUCESSO Then gError 100002

    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        For Each objRastroEstIni In colRastreamento
    
            Select Case objRastroEstIni.iEscaninho
        
                Case ESCANINHO_DISPONIVEL
                    dQuantDisp = dQuantDisp + objRastroEstIni.dQuantidade
                    
                Case ESCANINHO_CONSERTO_NOSSO
                    dQuantConserto = dQuantConserto + objRastroEstIni.dQuantidade
                
                Case ESCANINHO_CONSIG_NOSSO
                    dQuantConsig = dQuantConsig + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_DEMO_NOSSO
                    dQuantDemo = dQuantDemo + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_OUTROS_NOSSO
                    dQuantOutros = dQuantOutros + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_BENEF_NOSSO
                    dQuantBenef = dQuantBenef + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_CONSERTO_3
                    dQuantConserto3 = dQuantConserto3 + objRastroEstIni.dQuantidade
                
                Case ESCANINHO_CONSIG_3
                    dQuantConsig3 = dQuantConsig3 + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_DEMO_3
                    dQuantDemo3 = dQuantDemo3 + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_OUTROS_3
                    dQuantOutros3 = dQuantOutros3 + objRastroEstIni.dQuantidade
        
                Case ESCANINHO_BENEF_3
                    dQuantBenef3 = dQuantBenef3 + objRastroEstIni.dQuantidade
    
            End Select
    
        Next
    
        If dQuantDisp <> objEstoqueProduto.dQuantidadeInicial Then gError 83461
        
        If dQuantConserto <> objEstoqueProduto.dQuantInicialConserto Then gError 83462
        
        If dQuantConsig <> objEstoqueProduto.dQuantInicialConsig Then gError 83463
        
        If dQuantDemo <> objEstoqueProduto.dQuantInicialDemo Then gError 83464
        
        If dQuantOutros <> objEstoqueProduto.dQuantInicialOutras Then gError 83465
        
        If dQuantBenef <> objEstoqueProduto.dQuantInicialBenef Then gError 83466
        
        If dQuantConserto3 <> objEstoqueProduto.dQuantInicialConserto3 Then gError 83467
        
        If dQuantConsig3 <> objEstoqueProduto.dQuantInicialConsig3 Then gError 83468
        
        If dQuantDemo3 <> objEstoqueProduto.dQuantInicialDemo3 Then gError 83469
        
        If dQuantOutros3 <> objEstoqueProduto.dQuantInicialOutras3 Then gError 83470
        
        If dQuantBenef3 <> objEstoqueProduto.dQuantInicialBenef3 Then gError 83471
    
    End If
    
    Testa_Rastreamento = SUCESSO
    
    Exit Function
    
Erro_Testa_Rastreamento:

    Testa_Rastreamento = gErr

    Select Case gErr
    
        Case 100001, 100002
        
        Case 83461
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_DISP", gErr, objEstoqueProduto.dQuantidadeInicial, dQuantDisp)
    
        Case 83462
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_CONSERTO", gErr, objEstoqueProduto.dQuantInicialConserto, dQuantConserto)
    
        Case 83463
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_CONSIG", gErr, objEstoqueProduto.dQuantInicialConsig, dQuantConsig)
    
        Case 83464
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_DEMO", gErr, objEstoqueProduto.dQuantInicialDemo, dQuantDemo)
    
        Case 83465
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_OUTROS", gErr, objEstoqueProduto.dQuantInicialOutras, dQuantOutros)
    
        Case 83466
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_BENEF", gErr, objEstoqueProduto.dQuantInicialBenef, dQuantBenef)
    
        Case 83467
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_CONSERTO3", gErr, objEstoqueProduto.dQuantInicialConserto3, dQuantConserto3)
    
        Case 83468
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_CONSIG3", gErr, objEstoqueProduto.dQuantInicialConsig3, dQuantConsig3)
    
        Case 83469
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_DEMO3", gErr, objEstoqueProduto.dQuantInicialDemo3, dQuantDemo3)
    
        Case 83470
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_OUTROS3", gErr, objEstoqueProduto.dQuantInicialOutras3, dQuantOutros3)
    
        Case 83471
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTO_DIF_INICIAL_BENEF3", gErr, objEstoqueProduto.dQuantInicialBenef3, dQuantBenef3)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155759)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objEstoqueProduto As ClassEstoqueProduto) As Long
'Move os dados da Tela para objEstoqueProduto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 27202
        
        objEstoqueProduto.sProduto = sProdutoFormatado

    End If
    
    lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
    If lErro <> SUCESSO Then Error 44928
    
    'Preenche objEstoqueProduto
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    objEstoqueProduto.sLocalizacaoFisica = LocalizacaoFisica.Text
    
    'Verifica se o modulo de contabilidade esta ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 22161

            'conta não cadastrada
            If lErro = 5700 Then Error 44938
    
            objEstoqueProduto.sContaContabil = sContaFormatada
            
        End If
        
    Else
        objEstoqueProduto.sContaContabil = ""
        
    End If
    
    If Len(Trim(Quantidade.ClipText)) > 0 Then
        objEstoqueProduto.dQuantidadeInicial = CDbl(Quantidade.Text)
    End If

    If Len(Trim(Valor.Text)) > 0 Then
        objEstoqueProduto.dSaldoInicial = CDbl(Valor.Text)
    End If

    If Len(Trim(DataInicial.ClipText)) > 0 Then objEstoqueProduto.dtDataInicial = CDate(DataInicial.Text)

    'Consig
    If Len(Trim(QuantConsig.ClipText)) > 0 Then
        objEstoqueProduto.dQuantInicialConsig = CDbl(QuantConsig.Text)
    End If
    If Len(Trim(ValorConsig.Text)) > 0 Then
        objEstoqueProduto.dValorInicialConsig = CDbl(ValorConsig.Text)
    End If
    If Len(Trim(QuantConsig3.ClipText)) > 0 Then
        objEstoqueProduto.dQuantInicialConsig3 = CDbl(QuantConsig3.Text)
    End If
    If Len(Trim(ValorConsig3.Text)) > 0 Then
        objEstoqueProduto.dValorInicialConsig3 = CDbl(ValorConsig3.Text)
    End If

    'Outras
    If Len(Trim(QuantOutras.ClipText)) > 0 Then
        objEstoqueProduto.dQuantInicialOutras = CDbl(QuantOutras.Text)
    End If
    If Len(Trim(ValorOutras.Text)) > 0 Then
        objEstoqueProduto.dValorInicialOutras = CDbl(ValorOutras.Text)
    End If
    If Len(Trim(QuantOutras3.ClipText)) > 0 Then
        objEstoqueProduto.dQuantInicialOutras3 = CDbl(QuantOutras3.Text)
    End If
    If Len(Trim(ValorOutras3.Text)) > 0 Then
        objEstoqueProduto.dValorInicialOutras3 = CDbl(ValorOutras3.Text)
    End If

'    If giTipoVersao = VERSAO_FULL Then
    
        'Conserto
        If Len(Trim(QuantConserto.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialConserto = CDbl(QuantConserto.Text)
        End If
        If Len(Trim(ValorConserto.Text)) > 0 Then
            objEstoqueProduto.dValorInicialConserto = CDbl(ValorConserto.Text)
        End If
        If Len(Trim(QuantConserto3.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialConserto3 = CDbl(QuantConserto3.Text)
        End If
        If Len(Trim(ValorConserto3.Text)) > 0 Then
            objEstoqueProduto.dValorInicialConserto3 = CDbl(ValorConserto3.Text)
        End If
    
        'Demo
        If Len(Trim(QuantDemo.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialDemo = CDbl(QuantDemo.Text)
        End If
        If Len(Trim(ValorDemo.Text)) > 0 Then
            objEstoqueProduto.dValorInicialDemo = CDbl(ValorDemo.Text)
        End If
        If Len(Trim(QuantDemo3.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialDemo3 = CDbl(QuantDemo3.Text)
        End If
        If Len(Trim(ValorDemo3.Text)) > 0 Then
            objEstoqueProduto.dValorInicialDemo3 = CDbl(ValorDemo3.Text)
        End If
    
    
        'Benef
        If Len(Trim(QuantBenef.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialBenef = CDbl(QuantBenef.Text)
        End If
        If Len(Trim(ValorBenef.Text)) > 0 Then
            objEstoqueProduto.dValorInicialBenef = CDbl(ValorBenef.Text)
        End If
        If Len(Trim(QuantBenef3.ClipText)) > 0 Then
            objEstoqueProduto.dQuantInicialBenef3 = CDbl(QuantBenef3.Text)
        End If
        If Len(Trim(ValorBenef3.Text)) > 0 Then
            objEstoqueProduto.dValorInicialBenef3 = CDbl(ValorBenef3.Text)
        End If
    
'    End If
    
    objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 27202, 44928

        Case 22161
            ContaContabil.SetFocus
        
        Case 44938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
            ContaContabil.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155760)

    End Select

    Exit Function

End Function

Public Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataInicial.SetFocus

    If Len(Trim(DataInicial.ClipText)) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 21327

        DataInicial.Text = Format(CDate(sData), "dd/mm/yy")

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 21327

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155761)

    End Select

    Exit Sub

End Sub

Public Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 21328

        DataInicial.Text = Format(CDate(sData), "dd/mm/yy")

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 21328

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155762)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objSldMesEst As New ClassSldMesEst
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iMes As Integer

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 21309

    End If
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155763)

    End Select

    Exit Sub

End Sub

Private Function Trata_Troca_Produto_Almoxarifado() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iAlmoxarifado As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objProdutoFilial As New ClassProdutoFilial
Dim vbMsgRes As VbMsgBoxResult
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Trata_Troca_Produto_Almoxarifado

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 33647

    If lErro = 25041 Then Error 33641

    lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
    If lErro <> SUCESSO Then Error 44929
    
    iAlmoxarifado = objAlmoxarifado.iCodigo

    lErro = EstoqueInicial_TrazTela(objProduto, iAlmoxarifado, gcolRastreamento)
    If lErro <> SUCESSO Then Error 44924

    'Preenche objProdutoFilial
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    objProdutoFilial.sProduto = objProduto.sCodigo
    
    'Lê ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then Error 33649

    If lErro = SUCESSO Then
        If objProdutoFilial.iAlmoxarifado = iAlmoxarifado Then
            Padrao.Value = MARCADO
        Else
            Padrao.Value = DESMARCADO
        End If
    End If



    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 45920

    Trata_Troca_Produto_Almoxarifado = SUCESSO
    
    Exit Function
    
Erro_Trata_Troca_Produto_Almoxarifado:

    Trata_Troca_Produto_Almoxarifado = Err
    
    Select Case Err
    
        Case 33647
            Produto.SetFocus

        Case 33641
            'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                DescricaoProduto.Caption = ""
                UnidMed.Caption = ""
                Call CF2(Me, "EstIni_Produto_ExibeInfo", "")
                'Segura o foco
                Produto.SetFocus
            End If

        Case 33649, 44924, 44929, 45920

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155764)
            
    End Select
    
    Exit Function

End Function

Public Sub ListaAlmoxarifado_DblClick()

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaEnxuta As String

On Error GoTo Erro_ListaAlmoxarifado_DblClick

    'Guarda o valor do código do Almoxarifado selecionada na ListBox AlmoxarifadoList
    objAlmoxarifado.iCodigo = ListaAlmoxarifado.ItemData(ListaAlmoxarifado.ListIndex)

    'Lê o Almoxarifado no BD
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then Error 44917

    Almoxarifado.Text = objAlmoxarifado.sNomeReduzido

    'Se Almoxarifado não está cadastrado, erro
    If lErro = 25056 Then Error 44918
    
    'preenche a conta contabil
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
                
        If Len(Trim(objAlmoxarifado.sContaContabil)) > 0 Then
            sContaEnxuta = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objAlmoxarifado.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61371
            
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
        Else
            ContaContabil.PromptInclude = False
            ContaContabil.Text = ""
            ContaContabil.PromptInclude = True
        End If
        
    End If

    
    Exit Sub

Erro_ListaAlmoxarifado_DblClick:

    Select Case Err

        Case 44917
    
        Case 44918
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
            ListaAlmoxarifado.RemoveItem (ListaAlmoxarifado.ListIndex)
    
        Case 61371
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objAlmoxarifado.sContaContabil)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155765)

    End Select

    Exit Sub

End Sub

Function Carga_ListaAlmoxarifado() As Long

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim colAlmoxarifados As New Collection

On Error GoTo Erro_Carga_ListaAlmoxarifado

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 44920

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        ListaAlmoxarifado.AddItem objAlmoxarifado.sNomeReduzido
        ListaAlmoxarifado.ItemData(ListaAlmoxarifado.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carga_ListaAlmoxarifado = SUCESSO
    
    Exit Function
    
Erro_Carga_ListaAlmoxarifado:

    Carga_ListaAlmoxarifado = Err

    Select Case Err
    
        Case 44920

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155766)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ESTOQUE_INICIAL
    Set Form_Load_Ocx = Me
    Caption = "Estoque Inicial"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "EstoqueInicial"
    
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

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel1_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call AlmoxarifadoLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        End If
    End If

End Sub

'Alteracoes Raphael 21/07/00

Public Sub QuantConserto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantConserto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantConserto_Validate

    If Len(Trim(QuantConserto.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantConserto.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantConserto.Text, ESCANINHO_CONSERTO_NOSSO, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 83004

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 83005

    Exit Sub

Erro_QuantConserto_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 83004

        Case 83005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantConserto.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155767)

    End Select

    Exit Sub

End Sub

Public Sub ValorConserto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorConserto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorConserto_Validate

    If Len(Trim(ValorConserto.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorConserto.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorConserto_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155768)

    End Select

    Exit Sub

End Sub

Public Sub QuantConsig_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantConsig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantConsig_Validate

    If Len(Trim(QuantConsig.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantConsig.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantConsig.Text, ESCANINHO_CONSIG_NOSSO, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 83002

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 83003

    Exit Sub

Erro_QuantConsig_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 83002

        Case 83003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantConsig.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155769)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorConsig_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorConsig_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorConsig_Validate

    If Len(Trim(ValorConsig.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorConsig.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorConsig_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155770)

    End Select

    Exit Sub

End Sub

Public Sub QuantDemo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantDemo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantDemo_Validate

    If Len(Trim(QuantDemo.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantDemo.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantDemo.Text, ESCANINHO_DEMO_NOSSO, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 83000

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 83001

    Exit Sub

Erro_QuantDemo_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 83000

        Case 83001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantDemo.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155771)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorDemo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDemo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDemo_Validate

    If Len(Trim(ValorDemo.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorDemo.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorDemo_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155772)

    End Select

    Exit Sub

End Sub

Public Sub QuantOutras_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantOutras_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantOutras_Validate

    If Len(Trim(QuantOutras.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantOutras.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantOutras.Text, ESCANINHO_OUTROS_NOSSO, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71998

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71999

    Exit Sub

Erro_QuantOutras_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71998

        Case 71999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantOutras.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155773)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorOutras_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorOutras_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorOutras_Validate

    If Len(Trim(ValorOutras.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorOutras.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorOutras_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155774)

    End Select

    Exit Sub

End Sub

Public Sub QuantBenef_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantBenef_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantBenef_Validate

    If Len(Trim(QuantBenef.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantBenef.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantBenef.Text, ESCANINHO_BENEF_NOSSO, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71996

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71997

    Exit Sub

Erro_QuantBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71996

        Case 71997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantBenef.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155775)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorBenef_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorBenef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorBenef_Validate

    If Len(Trim(ValorBenef.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorBenef.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorBenef_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155776)

    End Select

    Exit Sub

End Sub

Public Sub QuantConserto3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantConserto3_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantConserto3_Validate

    If Len(Trim(QuantConserto3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantConserto3.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantConsig3.Text, ESCANINHO_CONSERTO_3, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71994

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71995

    Exit Sub

Erro_QuantConserto3_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71994

        Case 71995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantConserto3.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155777)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorConserto3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorConserto3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorConserto3_Validate

    If Len(Trim(ValorConserto3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorConserto3.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorConserto3_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155778)

    End Select

    Exit Sub

End Sub

Public Sub QuantConsig3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantConsig3_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantConsig3_Validate

    If Len(Trim(QuantConsig3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantConsig3.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantConsig3.Text, ESCANINHO_CONSIG_3, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71992

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71993

    Exit Sub

Erro_QuantConsig3_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71992

        Case 71993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantConsig3.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155779)

    End Select
    
    Exit Sub

End Sub

Public Sub ValorConsig3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorConsig3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorConsig3_Validate

    If Len(Trim(ValorConsig3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorConsig3.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorConsig3_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155780)

    End Select

    Exit Sub

End Sub

Public Sub QuantDemo3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantDemo3_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantDemo3_Validate

    If Len(Trim(QuantDemo3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantDemo3.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantDemo3.Text, ESCANINHO_DEMO_3, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71990

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71991

    Exit Sub

Erro_QuantDemo3_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71990

        Case 71991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantDemo3.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155781)

    End Select

    Exit Sub

End Sub

Public Sub ValorDemo3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDemo3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDemo3_Validate

    If Len(Trim(ValorDemo3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorDemo3.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorDemo3_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155782)

    End Select

    Exit Sub

End Sub

Public Sub QuantOutras3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantOutras3_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantOutras3_Validate

    If Len(Trim(QuantOutras3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantOutras3.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantOutras3.Text, ESCANINHO_OUTROS_3, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71988

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71989

    Exit Sub

Erro_QuantOutras3_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71988

        Case 71989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantOutras3.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155783)

    End Select

    Exit Sub

End Sub

Public Sub ValorOutras3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorOutras3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorOutras3_Validate

    If Len(Trim(ValorOutras3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorOutras3.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorOutras3_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155784)

    End Select

    Exit Sub

End Sub

Public Sub QuantBenef3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantBenef3_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQuantRastro As Double

On Error GoTo Erro_QuantBenef3_Validate

    If Len(Trim(QuantBenef3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(QuantBenef3.Text)
        If lErro <> SUCESSO Then gError 33636

    End If

    'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.
    lErro = Rastreamento_Quantidade_Teste(gcolRastreamento, QuantBenef3.Text, ESCANINHO_BENEF_3, dQuantRastro)
    If lErro <> SUCESSO And lErro <> 71986 Then gError 71985

    'a quantidade do rastreamento é maior que a quantidade do estoque inicial
    If lErro = 71986 Then gError 71987

    Exit Sub

Erro_QuantBenef3_Validate:

    Cancel = True

    Select Case gErr

        Case 33636

        Case 71985

        Case 71987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTESTINI_MENOR_RASTREAMENTO", gErr, StrParaDbl(QuantBenef3.Text), dQuantRastro)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155785)

    End Select

    Exit Sub

End Sub

Public Sub ValorBenef3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorBenef3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorBenef3_Validate

    If Len(Trim(ValorBenef3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorBenef3.Text)
        If lErro <> SUCESSO Then Error 21309

    End If

    Exit Sub

Erro_ValorBenef3_Validate:

    Cancel = True

    Select Case Err

        Case 21309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155786)

    End Select

    Exit Sub

End Sub

Private Function Rastreamento_Quantidade_Teste(colRastreamento As Collection, sQuantidade As String, ByVal iEscaninho As Integer, dQuantRastro As Double) As Long
'testa os rastreamentos deste escaninho contra a quantidade do mesmo. Se for maior ==> erro.

Dim objRastroEstIni As ClassRastroEstIni
Dim lErro As Long
Dim dQuantEstIni As Double

On Error GoTo Erro_Rastreamento_Quantidade_Teste

    If Not (colRastreamento Is Nothing) Then
    
        dQuantEstIni = StrParaDbl(sQuantidade)
    
        For Each objRastroEstIni In colRastreamento
            
            If objRastroEstIni.iEscaninho = iEscaninho Then
            
                dQuantRastro = dQuantRastro + objRastroEstIni.dQuantidade
                
            End If
            
        Next
            
        If dQuantEstIni < dQuantRastro Then gError 71986

    End If

    Rastreamento_Quantidade_Teste = SUCESSO
    
    Exit Function
    
Erro_Rastreamento_Quantidade_Teste:

    Rastreamento_Quantidade_Teste = gErr
    
    Select Case gErr
    
        Case 71986
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155787)
        
    End Select
        
    Exit Function

End Function

Public Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Se o Frame selecionado for o frame 2
        If TabStrip1.SelectedItem.Index = TAB_COMPLEMENTAR Then
            
            'preenche os labels
            AlmoxarifadoComplementar.Caption = Almoxarifado.Text
            ProdutoComplementar.Caption = Produto.Text
            UMComplementar = UnidMed.Caption
        
        End If
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

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

Function EstIni_Produto_ExibeInfo(ByVal objCT As Object, ByVal sProduto As String, Optional ByVal sCor As String = "") As Long
    EstIni_Produto_ExibeInfo = SUCESSO
End Function

