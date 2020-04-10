VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTRastreamentoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

Dim iFrameAtual As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variável de controle de alteração no campo produto
Dim iProdutoAlterado As Integer
Dim iClienteAlterado As Integer

Dim objGrid As AdmGrid

'Variáveis que indicam os índices de cada coluna do grid
Dim iGridTestes_Teste_Col As Integer
Dim iGridTestes_ResultadoNaoConforme_Col As Integer
Dim iGridTestes_ResultadoValor_Col As Integer
Dim iGridTestes_ResultadoObs_Col As Integer
Dim iGridTestes_RegistroAnaliseID_Col As Integer
Dim iGridTestes_RegistroAnaliseData_Col As Integer
Dim iGridTestes_LimiteDe_Col As Integer
Dim iGridTestes_LimiteAte_Col As Integer
Dim iGridTestes_NoCertificado_Col As Integer
Dim iGridTestes_Especificacao_Col As Integer
Dim iGridTestes_Observacao_Col As Integer
Dim iGridTestes_Metodo_Col As Integer

'Variáveis Globais
Public iAlterado As Integer

'Eventos dos Browses
Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Dim WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
'Declaração utilizada para evento BotaoTestes_Click
Dim WithEvents objEventoTestes As AdmEvento
Attribute objEventoTestes.VB_VarHelpID = -1
Dim WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Localizacao() As Object
     Set Localizacao = objUserControl.Controls("Localizacao")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get FilialCli() As Object
     Set FilialCli = objUserControl.Controls("FilialCli")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get IdAnalisePadrao() As Object
     Set IdAnalisePadrao = objUserControl.Controls("IdAnalisePadrao")
End Property

Public Property Get BotaoTrazerTeste() As Object
     Set BotaoTrazerTeste = objUserControl.Controls("BotaoTrazerTeste")
End Property

Public Property Get BotaoImprimirLaudo() As Object
     Set BotaoImprimirLaudo = objUserControl.Controls("BotaoImprimirLaudo")
End Property

Public Property Get BotaoTestes() As Object
     Set BotaoTestes = objUserControl.Controls("BotaoTestes")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get LabelObservacao() As Object
     Set LabelObservacao = objUserControl.Controls("LabelObservacao")
End Property

Public Property Get LabelEspecificacao() As Object
     Set LabelEspecificacao = objUserControl.Controls("LabelEspecificacao")
End Property

Public Property Get NoCertificado2() As Object
     Set NoCertificado2 = objUserControl.Controls("NoCertificado2")
End Property

Public Property Get FrameLimites() As Object
     Set FrameLimites = objUserControl.Controls("FrameLimites")
End Property

Public Property Get LabelLimiteAte() As Object
     Set LabelLimiteAte = objUserControl.Controls("LabelLimiteAte")
End Property

Public Property Get LabelLimiteDe() As Object
     Set LabelLimiteDe = objUserControl.Controls("LabelLimiteDe")
End Property

Public Property Get LabelMetodo() As Object
     Set LabelMetodo = objUserControl.Controls("LabelMetodo")
End Property

Public Property Get LabelTeste() As Object
     Set LabelTeste = objUserControl.Controls("LabelTeste")
End Property

Public Property Get LabelTesteCodigo() As Object
     Set LabelTesteCodigo = objUserControl.Controls("LabelTesteCodigo")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get ResultadoObs() As Object
     Set ResultadoObs = objUserControl.Controls("ResultadoObs")
End Property

Public Property Get Especificacao() As Object
     Set Especificacao = objUserControl.Controls("Especificacao")
End Property

Public Property Get ObservacaoTeste() As Object
     Set ObservacaoTeste = objUserControl.Controls("ObservacaoTeste")
End Property

Public Property Get NoCertificado() As Object
     Set NoCertificado = objUserControl.Controls("NoCertificado")
End Property

Public Property Get Metodo() As Object
     Set Metodo = objUserControl.Controls("Metodo")
End Property

Public Property Get RegistroAnaliseData() As Object
     Set RegistroAnaliseData = objUserControl.Controls("RegistroAnaliseData")
End Property

Public Property Get RegistroAnaliseID() As Object
     Set RegistroAnaliseID = objUserControl.Controls("RegistroAnaliseID")
End Property

Public Property Get ResultadoNaoConforme() As Object
     Set ResultadoNaoConforme = objUserControl.Controls("ResultadoNaoConforme")
End Property

Public Property Get ResultadoValor() As Object
     Set ResultadoValor = objUserControl.Controls("ResultadoValor")
End Property

Public Property Get Teste() As Object
     Set Teste = objUserControl.Controls("Teste")
End Property

Public Property Get LimiteDe() As Object
     Set LimiteDe = objUserControl.Controls("LimiteDe")
End Property

Public Property Get LimiteAte() As Object
     Set LimiteAte = objUserControl.Controls("LimiteAte")
End Property

Public Property Get GridTestes() As Object
     Set GridTestes = objUserControl.Controls("GridTestes")
End Property

Public Property Get FilialOP() As Object
     Set FilialOP = objUserControl.Controls("FilialOP")
End Property

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get LoteLabel() As Object
     Set LoteLabel = objUserControl.Controls("LoteLabel")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get UpDownValidade() As Object
     Set UpDownValidade = objUserControl.Controls("UpDownValidade")
End Property

Public Property Get DataValidade() As Object
     Set DataValidade = objUserControl.Controls("DataValidade")
End Property

Public Property Get UpDownFabricacao() As Object
     Set UpDownFabricacao = objUserControl.Controls("UpDownFabricacao")
End Property

Public Property Get DataFabricacao() As Object
     Set DataFabricacao = objUserControl.Controls("DataFabricacao")
End Property

Public Property Get UpDownEntrada() As Object
     Set UpDownEntrada = objUserControl.Controls("UpDownEntrada")
End Property

Public Property Get DataEntrada() As Object
     Set DataEntrada = objUserControl.Controls("DataEntrada")
End Property

Public Property Get BotaoRastroMovto() As Object
     Set BotaoRastroMovto = objUserControl.Controls("BotaoRastroMovto")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

'--- fim dos properties get dos controles da tela

Private Function Carrega_FilialOP() As Long
'Carrega FilialOP a partir das Filiais Empresas

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialOP

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then
        
            FilialOP.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialOP.ItemData(FilialOP.NewIndex) = objFiliais.iCodFilial
    
        End If
        
    Next
        
    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165877)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objRastroLote As ClassRastreamentoLote) As Long

Dim lErro As Long
Dim bCancel As Boolean
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se foi passado algum objRastroLote como parâmetro
    If Not (objRastroLote Is Nothing) Then

        'Lê RastremantoLote a partir do Produto, FIlialOP e Código do lote passados
        lErro = CF("RastreamentoLote_Le", objRastroLote)
        If lErro <> SUCESSO And lErro <> 75710 Then gError 75711

        'Se o RastreamentoLote está cadastrado
        If lErro <> 75710 Then

            'Traz RastreamentoLote para a tela
            lErro = Traz_RastroLote_Tela(objRastroLote)
            If lErro <> SUCESSO Then gError 75712

        'Se não está cadastrado
        Else

            'Exibe dados principais do rastreamento
            Lote.Text = objRastroLote.sCodigo

            Produto.PromptInclude = False
            Produto.Text = objRastroLote.sProduto
            Produto.PromptInclude = True

            'Dispara validate de Produto
            Call Produto_Validate(bCancel)

            If objRastroLote.iFilialOP <> 0 Then

                iAchou = 0
            
                'Se o Rastreamento possui FilialOP (Rastro Por Ordem de Produção)
                If objRastroLote.iFilialOP <> 0 Then
            
                    For iIndice = 0 To FilialOP.ListCount - 1
                        If FilialOP.ItemData(iIndice) = objRastroLote.iFilialOP Then
                            iAchou = 1
                            FilialOP.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                    
                    'se a filial não foi encontrada ==> erro
                    If iAchou = 0 Then gError 75718
            
                End If

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 75711, 75712, 75717

        Case 75718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRastroLote.iFilialOP)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165878)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RastreamentoLote"

    'Lê os atributos de objRastroLote que aparecem na Tela
    lErro = Move_Tela_Memoria(objRastroLote)
    If lErro <> SUCESSO Then gError 75735

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objRastroLote.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Lote", objRastroLote.sCodigo, STRING_LOTE_RASTREAMENTO, "Lote" 'Alterado por Wagner 09/03/2006
    colCampoValor.Add "FilialOP", objRastroLote.iFilialOP, 0, "FilialOP"
    colCampoValor.Add "Produto", objRastroLote.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "DataEntrada", objRastroLote.dtDataEntrada, 0, "DataEntrada"
    colCampoValor.Add "DataValidade", objRastroLote.dtDataValidade, 0, "DataValidade"
    colCampoValor.Add "DataFabricacao", objRastroLote.dtDataFabricacao, 0, "DataFabricacao"
    colCampoValor.Add "Observacao", objRastroLote.sObservacao, STRING_RASTRO_OBSERVACAO, "Observacao"

    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, RATREAMENTOLOTE_STATUS_BAIXADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 75735

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165879)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objRastroLote
    objRastroLote.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objRastroLote.sCodigo = colCampoValor.Item("Lote").vValor
    objRastroLote.iFilialOP = colCampoValor.Item("FilialOP").vValor
    objRastroLote.sProduto = colCampoValor.Item("Produto").vValor
    objRastroLote.dtDataEntrada = colCampoValor.Item("DataEntrada").vValor
    objRastroLote.dtDataValidade = colCampoValor.Item("DataValidade").vValor
    objRastroLote.dtDataFabricacao = colCampoValor.Item("DataFabricacao").vValor
    objRastroLote.sObservacao = colCampoValor.Item("Observacao").vValor

    'Traz dados do RastreamentoLote para a tela
    lErro = Traz_RastroLote_Tela(objRastroLote)
    If lErro <> SUCESSO Then gError 75736

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 75736

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165880)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objRastroLote As ClassRastreamentoLote) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    'Move dados da tela para a memória
    objRastroLote.sCodigo = Trim(Lote.Text)
    objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)

    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 75737

    objRastroLote.sProduto = sProdutoFormatado
    objRastroLote.dtDataEntrada = StrParaDate(DataEntrada.Text)
    objRastroLote.dtDataFabricacao = StrParaDate(DataFabricacao.Text)
    objRastroLote.dtDataValidade = StrParaDate(DataValidade.Text)
    objRastroLote.sObservacao = Observacao.Text
    objRastroLote.sLocalizacao = Localizacao.Text

    If Len(Trim(Cliente.Text)) > 0 Then
    
        'Lê o Cliente a partir do Nome Reduzido
        objCliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 106645
        
        'Se não econtrou o Cliente, erro
        If lErro = 12348 Then gError 106646
        
        objRastroLote.lCliente = objCliente.lCodigo
    
    End If
    
    If Len(Trim(FilialCli.Text)) > 0 Then objRastroLote.iFilialCli = Codigo_Extrai(FilialCli.Text)
    
    lErro = CF2(Me, "RL_Move_Tela_Memoria", objRastroLote)
    If lErro <> SUCESSO Then gError 75737
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Select Case gErr

        Case 75737, 106645

        Case 106646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165881)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "RL_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set objGrid = Nothing
     
    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoTestes = Nothing
    Set objEventoCliente = Nothing
    Set objEventoLocalizacao = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
   lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    Lote.MaxLength = STRING_LOTE_RASTREAMENTO

    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoTestes = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoLocalizacao = New AdmEvento

    'Carrega FilialOP a partir das Filiais Empresas
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 83009

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 75714

    lErro = Inicializa_GridTestes
    If lErro <> SUCESSO Then gError 75714
    
    iClienteAlterado = 0
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 75714, 83009

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165882)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Traz_RastroLote_Tela(objRastroLote As ClassRastreamentoLote) As Long
'Traz dados do RastreamentoLote para a tela

Dim lErro As Long
Dim bCancel As Boolean
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Traz_RastroLote_Tela

    'Limpa a tela
    Call Limpa_Tela_RastroLote

    'Lê RastremantoLote a partir do Produto, FIlialOP e Código do lote passados
    lErro = CF("RastreamentoLote_Le", objRastroLote)
    If lErro <> SUCESSO And lErro <> 75710 Then gError 130318
    
    lErro = CF("RastreamentoLote_Le_Testes", objRastroLote)
    If lErro <> SUCESSO Then gError 130318
    
    'Código do lote
    Lote.Text = objRastroLote.sCodigo

    iAchou = 0

    'Se o Rastreamento possui FilialOP (Rastro Por Ordem de Produção)
    If objRastroLote.iFilialOP <> 0 Then

        For iIndice = 0 To FilialOP.ListCount - 1
            If FilialOP.ItemData(iIndice) = objRastroLote.iFilialOP Then
                iAchou = 1
                FilialOP.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'se a filial não foi encontrada ==> erro
        If iAchou = 0 Then gError 75723

    End If
    
    'Produto
    Produto.PromptInclude = False
    Produto.Text = objRastroLote.sProduto
    Produto.PromptInclude = True

    'Dispara o validate de produto
    Call Produto_Validate(bCancel)

    Call DateParaMasked(DataEntrada, objRastroLote.dtDataEntrada)
    Call DateParaMasked(DataValidade, objRastroLote.dtDataValidade)
    Call DateParaMasked(DataFabricacao, objRastroLote.dtDataFabricacao)

    Observacao.Text = objRastroLote.sObservacao
    
    Localizacao.Text = objRastroLote.sLocalizacao

    If objRastroLote.lCliente <> 0 Then
    
        Cliente.Text = objRastroLote.lCliente
        Call Cliente_Validate(bSGECancelDummy)
        
        FilialCli.Text = objRastroLote.iFilialCli
        Call FilialCli_Validate(bSGECancelDummy)
    
    End If
    
    lErro = Traz_Qualidade_Tela(objRastroLote)
    If lErro <> SUCESSO Then gError 130316
    
    lErro = CF2(Me, "RL_Traz_RastroLote_Tela", objRastroLote)
    If lErro <> SUCESSO Then gError 130316
    
    iClienteAlterado = 0
    iAlterado = 0

    Traz_RastroLote_Tela = SUCESSO

    Exit Function

Erro_Traz_RastroLote_Tela:

    Select Case gErr

        Case 75722, 130316, 130318

        Case 75723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRastroLote.iFilialOP)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165883)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o Produto Foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 75766

    'Verifica se o Lote foi preenchido
    If Len(Trim(Lote.Text)) = 0 Then gError 75767

    'Não pode excluir Lote vinculado a OP
    If Len(Trim(FilialOP.Text)) > 0 Then gError 75904

    'Formata o Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 75768

    objRastroLote.sProduto = sProdutoFormatado
    objRastroLote.sCodigo = Lote.Text
    
    'Verifica se RastreamentoLote existe
    lErro = CF("RastreamentoLote_Le", objRastroLote)
    If lErro <> SUCESSO And lErro <> 75710 Then gError 75769

    'Se o RastreamentoLote não estiver cadastrado, erro
    If lErro = 75710 Then gError 75770

    'Pergunta se o usuário confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RASTREAMENTOLOTE", objRastroLote.sProduto, objRastroLote.sCodigo)
    If vbMsgRes = vbNo Then gError 75771

    'Exclui RastreamentoLote
    lErro = CF("RastreamentoLote_Exclui", objRastroLote)
    If lErro <> SUCESSO Then gError 75772

    'Limpa a tela
    Call Limpa_Tela_RastroLote

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 75766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 75767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", gErr)

        Case 75768, 75769, 75771, 75772

        Case 75770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastroLote.sProduto, objRastroLote.sCodigo)
        
        Case 75904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RASTREAMENTOOP", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165884)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Inicia gravação de RastreamentoLote
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 75738

    'Limpa a tela
    Call Limpa_Tela_RastroLote

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75738

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165885)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Gravar_Registro

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 75739

    'Verifica se a data de entrada foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 75740

    'Verifica se o lot foi preenchido
    If Len(Trim(Lote.Text)) = 0 Then gError 75741

    'Se a data de validade e a data de fabricação forma preenchidas
    If Len(Trim(DataValidade.ClipText)) > 0 And Len(Trim(DataFabricacao.ClipText)) > 0 Then
        'Verifica se a data de validade é menor que a data de fabricação
        If StrParaDate(DataValidade.Text) < StrParaDate(DataFabricacao.Text) Then gError 75742
    End If
    
    'Critica Rastreamento
    lErro = Rastro_Critica()
    If lErro <> SUCESSO Then gError 75905
    
    'Move dados da tela para a memória
    lErro = Move_Tela_Memoria(objRastroLote)
    If lErro <> SUCESSO Then gError 75724

    lErro = Move_Testes_Memoria(objRastroLote)
    If lErro <> SUCESSO Then gError 75724

    lErro = Trata_Alteracao(objRastroLote, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> SUCESSO Then Error 32283

    'Grava RastreamentoLote
    lErro = CF("RastreamentoLote_Grava", objRastroLote)
    If lErro <> SUCESSO Then gError 75743

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32283

        Case 75739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 75740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 75741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", gErr)

        Case 75742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_MENOR_DATAFABRICACAO", gErr)

        Case 75724, 75743, 75905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165886)

    End Select

    Exit Function

End Function

Public Sub BotaoImprimirLaudo_Click()
    
Dim lErro As Long, objOP As New ClassOrdemDeProducao, objRelatorio As New AdmRelatorio
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoImprimirLaudo_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 75739
    If Len(Trim(Lote.Text)) = 0 Then gError 130334
    
    objOP.sCodigo = Lote.Text
    
    If InStr(UCase(gsNomeEmpresa), "INPAL") <> 0 Then
        Call objRelatorio.Rel_Menu_Executar("Laudo do Controle de Qualidade de OP", objOP)
    Else
    
        'Formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 75737
            
        Call objRelatorio.ExecutarDireto("Laudo do Controle de Qualidade de OP", "", 1, "", "TLOTE", Lote.Text, "TPRODUTO", sProdutoFormatado)
        
    End If
    
    Exit Sub
     
Erro_BotaoImprimirLaudo_Click:

    Select Case gErr
    
        Case 75737
          
        Case 75739
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
          
        Case 130334
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165887)
     
    End Select
     
    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

    'Limpa a tela
    Call Limpa_Tela_RastroLote

End Sub

Sub Limpa_Tela_RastroLote()

    'Limpa a tela
    Call Limpa_Tela(Me)

    Descricao.Caption = ""
    FilialOP.ListIndex = -1

    FilialCli.Clear

    Call Grid_Limpa(objGrid)
    
    Call Teste_LimpaInfo
    
    Call CF2(Me, "RL_Limpa_Tela_RastroLote")
    
    iProdutoAlterado = 0
    iClienteAlterado = 0
    iAlterado = 0

End Sub

Private Function Rastro_Critica() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastroLote As New ClassRastreamentoLote
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Rastro_Critica
        
    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 75896
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 75894
     
    'Se não encontrou o produto
    If lErro = 28030 Then gError 75895
    
    'se o produto não estiver trabalhando com rastro ou rastro numero de serie ==> erro
    If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then gError 83011 ' Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then gError 83011 'Alterado por Wagner 09/03/2006
    
    'Se o Rastreamento do produto for por Lote e a Filial estiver preenchida ==> erro
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE And Len(Trim(FilialOP.Text)) <> 0 Then gError 83012
    
    '#######################################################################
    'Inserido por Wagner 09/03/2006
    'Se o Rastreamento do produto for por serie e a Filial estiver preenchida ==> erro
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE And Len(Trim(FilialOP.Text)) <> 0 Then gError 83012
    '#######################################################################
    
    'Se o Rastreamento do produto for por OP e a filial não estiver preenchida ==> erro
    If objProduto.iRastro = PRODUTO_RASTRO_OP And Len(Trim(FilialOP.Text)) = 0 Then gError 83014
    
    Exit Function
    
    Rastro_Critica = SUCESSO
    
Erro_Rastro_Critica:

    Rastro_Critica = gErr
    
    Select Case gErr
        
        Case 75894, 75896
        
        Case 75895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 83011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, objProduto.sCodigo)
        
        Case 83012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTROOP", gErr, objProduto.sCodigo)
        
        Case 83014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_RASTROOP_FILIAL_ZERADA", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165888)
    
    End Select
    
    Exit Function

End Function

Public Sub BotaoRastroMovto_Click()

Dim colSelecao As New Collection
Dim objRastroMovEstoque As New ClassRastroMovEstoque
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoRastroMovto_Click

    objRastroMovEstoque.sLote = Lote.Text
    
    objRastroMovEstoque.iFilialOP = Codigo_Extrai(FilialOP.Text)

    objRastroMovEstoque.dtData = StrParaDate(DataEntrada.Text)
    
    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83479

    objRastroMovEstoque.sProduto = sProdutoFormatado
    
    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroMovEstoqueLista", colSelecao, objRastroMovEstoque)

    Exit Sub

Erro_BotaoRastroMovto_Click:

    Select Case gErr

        Case 83479

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165889)

    End Select

    Exit Sub

End Sub


Public Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(DataValidade.ClipText) = 0 Then Exit Sub

    'Critica a data
    lErro = Data_Critica(DataValidade.Text)
    If lErro <> SUCESSO Then gError 75726

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True

    Select Case gErr

        Case 75726

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165890)

    End Select

    Exit Sub

End Sub

Public Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(DataEntrada.ClipText) = 0 Then Exit Sub

    'Critica a data
    lErro = Data_Critica(DataEntrada.Text)
    If lErro <> SUCESSO Then gError 75727

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 75727

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165891)

    End Select

    Exit Sub

End Sub

Public Sub DataFabricacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFabricacao_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(DataFabricacao.ClipText) = 0 Then Exit Sub

    'Critica a data
    lErro = Data_Critica(DataFabricacao.Text)
    If lErro <> SUCESSO Then gError 75728

    Exit Sub

Erro_DataFabricacao_Validate:

    Cancel = True

    Select Case gErr

        Case 75728

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165892)

    End Select

    Exit Sub

End Sub

Public Sub FilialOP_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOP_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialOP.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 83010

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 83011

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 83012

    Exit Sub

Erro_FilialOP_Validate:

    Cancel = True

    Select Case gErr

        Case 83010

        Case 83011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 83012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165893)

    End Select

    Exit Sub

End Sub

Public Sub FilialOP_Click()

Dim lErro As Long

On Error GoTo Erro_FilialOP_Click

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_FilialOP_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165894)

    End Select

    Exit Sub

End Sub

Private Sub Text1_Change()

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

End Sub

Public Sub UpDownValidade_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_DownClick

    'Verifica se a Data da Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataValidade, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75729

    Exit Sub

Erro_UpDownValidade_DownClick:

    Select Case gErr

        Case 75729

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165895)

    End Select

    Exit Sub

End Sub

Public Sub UpDownValidade_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_UpClick

    'Verifica se a Data de Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = Data_Up_Down_Click(DataValidade, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75730

    Exit Sub

Erro_UpDownValidade_UpClick:

    Select Case gErr

        Case 75730

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165896)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_DownClick

    'Verifica se a Data da Entrada foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75731

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 75731

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165897)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_UpClick

    'Verifica se a Data de Entrada foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75732

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 75732

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165898)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFabricacao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFabricacao_DownClick

    'Verifica se a Data da Fabricacao foi preenchida
    If Len(Trim(DataFabricacao.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataFabricacao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75733

    Exit Sub

Erro_UpDownFabricacao_DownClick:

    Select Case gErr

        Case 75733

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165899)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFabricacao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFabricacao_UpClick

    'Verifica se a Data de Fabricacao foi preenchida
    If Len(Trim(DataFabricacao.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = lErro = Data_Up_Down_Click(DataFabricacao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75734

    Exit Sub

Erro_UpDownFabricacao_UpClick:

    Select Case gErr

        Case 75734

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165900)

    End Select

    Exit Sub

End Sub

Public Sub LoteLabel_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote

    objRastroLote.sCodigo = Lote.Text

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRastroLote = obj1

    'Traz RastreamentoLote para a tela
    lErro = Traz_RastroLote_Tela(objRastroLote)
    If lErro <> SUCESSO Then gError 75713

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 75713

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165901)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoLabel_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

    'Chama tela de Browse de Produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto

    Set objProduto = obj1

    'Coloca produto na tela
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True

    Me.Show

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Produto_Validate

    'Se nao houve alteracao
    If iProdutoAlterado = 0 Then Exit Sub
    
    'Se produto não estiver preenchido --> limpa descrição
    If Len(Trim(Produto.ClipText)) = 0 Then
        Descricao.Caption = ""

    Else

        'Critica o produto
        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 75719

        'Produto não cadastrado
        If lErro = 25041 Then gError 75720

        'se o produto não estiver trabalhando com rastro ou rastro numero de serie ==> erro
        If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then gError 83013 'Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then gError 83013 'Alterado por Wagner 09/03/2006
        
        Descricao.Caption = objProduto.sDescricao
        
        '??? se nao estiver carregando a tela e mudou de produto entao preencher o grid de testes com informacoes da tabela produtoteste
        
    End If

    iProdutoAlterado = 0
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 75719

        Case 75720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 83013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165902)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

     Unload Me

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = 1

End Sub

Public Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataFabricacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus()

    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Public Sub DataEntrada_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)

End Sub

Public Sub DataValidade_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)

End Sub

Public Sub DataFabricacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFabricacao, iAlterado)

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTREAMENTO_LOTE
    Set Form_Load_Ocx = Me
    Caption = "Cadastro de Lotes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastreamentoLote"

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

        If Me.ActiveControl Is Lote Then
            Call LoteLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If

        Call CF2(Me, "RL_UserControl_KeyDown", KeyCode, Shift)
        
    End If

End Sub

Private Function Inicializa_GridTestes() As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Teste")
    objGrid.colColuna.Add ("Não Conforme")
    objGrid.colColuna.Add ("Resultado")
    objGrid.colColuna.Add ("Anotação")
    objGrid.colColuna.Add ("Id. Análise")
    objGrid.colColuna.Add ("Data Análise")
    objGrid.colColuna.Add ("Limite De")
    objGrid.colColuna.Add ("Limite Até")
    objGrid.colColuna.Add ("Certificado")
    objGrid.colColuna.Add ("Especificação")
    objGrid.colColuna.Add ("Observação")
    objGrid.colColuna.Add ("Método")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Teste.Name)
    objGrid.colCampo.Add (ResultadoNaoConforme.Name)
    objGrid.colCampo.Add (ResultadoValor.Name)
    objGrid.colCampo.Add (ResultadoObs.Name)
    objGrid.colCampo.Add (RegistroAnaliseID.Name)
    objGrid.colCampo.Add (RegistroAnaliseData.Name)
    objGrid.colCampo.Add (LimiteDe.Name)
    objGrid.colCampo.Add (LimiteAte.Name)
    objGrid.colCampo.Add (NoCertificado.Name)
    objGrid.colCampo.Add (Especificacao.Name)
    objGrid.colCampo.Add (ObservacaoTeste.Name)
    objGrid.colCampo.Add (Metodo.Name)

    'Colunas do Grid
    iGridTestes_Teste_Col = 1
    iGridTestes_ResultadoNaoConforme_Col = 2
    iGridTestes_ResultadoValor_Col = 3
    iGridTestes_ResultadoObs_Col = 4
    iGridTestes_RegistroAnaliseID_Col = 5
    iGridTestes_RegistroAnaliseData_Col = 6
    iGridTestes_LimiteDe_Col = 7
    iGridTestes_LimiteAte_Col = 8
    iGridTestes_NoCertificado_Col = 9
    iGridTestes_Especificacao_Col = 10
    iGridTestes_Observacao_Col = 11
    iGridTestes_Metodo_Col = 12
    
    objGrid.objGrid = GridTestes

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_TESTES_PRODUTO + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridTestes.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridTestes = SUCESSO

End Function

Sub GridTestes_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If
    
End Sub

Sub GridTestes_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Sub GridTestes_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridTestes_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridTestes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridTestes_KeyDown

    'Guardo o item atual e o número de linhas existente
    iLinhasExistentesAnterior = objGrid.iLinhasExistentes
    iLinhaAnterior = GridTestes.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGrid)
    
    'se alguma linha foi excluída
    If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
        If objGrid.iLinhasExistentes = 0 Then
            Call Teste_LimpaInfo
        Else
            Call Teste_ExibeInfo(GridTestes.Row)
        End If
    End If
    
    Exit Sub

Erro_GridTestes_KeyDown:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165903)

    End Select

    Exit Sub

End Sub

Sub GridTestes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridTestes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Sub GridTestes_RowColChange()

    Call Grid_RowColChange(objGrid)

    If GridTestes.Row <> 0 Then
        Call Teste_ExibeInfo(GridTestes.Row)
    End If
    
End Sub

Sub GridTestes_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
        
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        
        'Teste
        Case Teste.Name
            If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) > 0 Then
                Teste.Enabled = False
            Else
                Teste.Enabled = True
            End If
        
        'Restante das colunas
        Case ResultadoNaoConforme.Name, ResultadoValor.Name, ResultadoObs.Name, RegistroAnaliseID.Name, RegistroAnaliseData.Name, LimiteDe.Name, LimiteAte.Name, NoCertificado.Name, Especificacao.Name, ObservacaoTeste.Name, Metodo.Name

            If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) = 0 Then
                
                ResultadoNaoConforme.Enabled = False
                ResultadoValor.Enabled = False
                ResultadoObs.Enabled = False
                RegistroAnaliseID.Enabled = False
                RegistroAnaliseData.Enabled = False
                LimiteDe.Enabled = False
                LimiteAte.Enabled = False
                NoCertificado.Enabled = False
                Especificacao.Enabled = False
                ObservacaoTeste.Enabled = False
                Metodo.Enabled = False

            Else
                
                ResultadoNaoConforme.Enabled = True
                ResultadoValor.Enabled = True
                ResultadoObs.Enabled = True
                RegistroAnaliseID.Enabled = True
                RegistroAnaliseData.Enabled = True
                LimiteDe.Enabled = True
                LimiteAte.Enabled = True
                NoCertificado.Enabled = True
                Especificacao.Enabled = True
                ObservacaoTeste.Enabled = True
                Metodo.Enabled = True

            End If
    
    End Select
    
    Call Teste_ExibeInfo(iLinha)
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165904)

    End Select

    Exit Sub

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridTestes.Col

            Case iGridTestes_Teste_Col

                lErro = Saida_Celula_Teste(objGridInt)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_ResultadoNaoConforme_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, ResultadoNaoConforme)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_ResultadoValor_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, ResultadoValor)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_ResultadoObs_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, ResultadoObs)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_RegistroAnaliseID_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, RegistroAnaliseID)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_RegistroAnaliseData_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, RegistroAnaliseData)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_LimiteDe_Col
                
                lErro = Saida_Celula_Limite(objGridInt, LimiteDe)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_LimiteAte_Col
            
                lErro = Saida_Celula_Limite(objGridInt, LimiteAte)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_NoCertificado_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, NoCertificado)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Especificacao_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, Especificacao)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Observacao_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, ObservacaoTeste)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Metodo_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, Metodo)
                If lErro <> SUCESSO Then gError 95069

        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 95076
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case 95068 To 95073, 95075, 95112
        
        Case 95076
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165905)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_Teste(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTeste As New ClassTestesQualidade
Dim iLinha As Integer
Dim sSiglaAux As Integer
Dim objTesteTextBox As Object
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Teste
    
    Set objGridInt.objControle = Teste

    'Verifica se a teste foi preenchida
    If Len(Trim(Teste.Text)) > 0 Then
        
        'Joga o controle Teste dentro de um obj
        Set objTesteTextBox = Teste
        
        'Le os dados da teste
        lErro = CF("TP_TesteQualidade_Le_Grid", objTesteTextBox, objTeste)
        If lErro <> SUCESSO Then gError 95077
        
        'Verifica se há alguma teste repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridTestes.Row Then
                                                    
                If GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido Then
                
                    Teste.Text = ""
                    gError 95081
                    
                End If
                    
            End If
                           
        Next
        
        lErro = Teste_Saida_Celula(GridTestes.Row, objTeste)
        If lErro <> SUCESSO Then gError 130337
    
    End If
             
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Teste = SUCESSO

    Exit Function

Erro_Saida_Celula_Teste:

    Saida_Celula_Teste = gErr

    Select Case gErr

        Case 95465
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTESQUALIDADE", gErr, objTeste.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95466
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTeste.sNomeReduzido)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Testes
                Call Chama_Tela("TestesQualidade", objTeste)
            Else
                Teste.SetFocus
            End If
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95081
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTEQUALIDADE_REPETIDO", gErr, objTeste.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95077, 95082, 130337
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165906)

    End Select

    Exit Function

End Function

Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Padrao
    
    Set objGridInt.objControle = objControle
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 95082
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165907)

    End Select

    Exit Function

End Function

Function Move_Testes_Memoria(objRastroLote As ClassRastreamentoLote) As Long

Dim lErro As Long, iLinha As Integer

On Error GoTo Erro_Move_Testes_Memoria

    'Para cada linha do grid
    For iLinha = 1 To objGrid.iLinhasExistentes
    
        'Verifica se a Teste foi preenchida
        If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) <> 0 Then
        
            lErro = Move_LinhaTeste_Memoria(objRastroLote, iLinha)
            If lErro <> SUCESSO Then gError 130315
            
        End If
                        
    Next
    
    Move_Testes_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_Testes_Memoria:

    Move_Testes_Memoria = gErr
     
    Select Case gErr
          
        Case 130315
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165908)
     
    End Select
     
    Exit Function

End Function

Function Move_LinhaTeste_Memoria(objRastroLote As ClassRastreamentoLote, ByVal iLinha As Integer) As Long

Dim lErro As Long, objResultadoTeste As ClassRastreamentoLoteTeste
Dim objTeste As New ClassTestesQualidade, dResultadoValor As Double

On Error GoTo Erro_Move_LinhaTeste_Memoria

    'Verifica se a Teste foi preenchida
    If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) <> 0 Then
                    
        Set objResultadoTeste = New ClassRastreamentoLoteTeste
        
        'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
        faz uma nova leitura em busca do codigo
        Set objTeste = New ClassTestesQualidade
        objTeste.sNomeReduzido = GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col)
        lErro = CF("TesteQualidade_Le_NomeReduzido", objTeste)
        If lErro <> SUCESSO And lErro <> 130109 Then gError 95462
        
        'Se nao achou => ERRO
        If lErro <> SUCESSO Then gError 95463
                                                                    
        'Preencher o obj
        With objResultadoTeste
        
            .iTesteCodigo = objTeste.iCodigo
            .iResultadoNaoConforme = StrParaInt(GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoNaoConforme_Col))
            .sResultadoValor = GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoValor_Col)
            .sResultadoObservacao = GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoObs_Col)
            .sRegistroAnaliseID = GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseID_Col)
            .dtRegistroAnaliseData = StrParaDate(GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseData_Col))
            .iSeqGrid = iLinha
            .sTesteEspecificacao = GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col)
            .iTesteTipoResultado = objTeste.iTipoResultado '??? nao aparece na tela
            .dTesteLimiteDe = StrParaDbl(GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col))
            .dTesteLimiteAte = StrParaDbl(GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col))
            .sTesteMetodoUsado = GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col)
            .sTesteObservacao = GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col)
            .iTesteNoCertificado = StrParaInt(GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col))
            
        End With
        
        objRastroLote.colTeste.Add objResultadoTeste

        If Len(Trim(objResultadoTeste.sResultadoValor)) = 0 Then gError 130341
        
        If objResultadoTeste.iTesteTipoResultado = 0 Then
        
            If Not (IsNumeric(objResultadoTeste.sResultadoValor)) Then gError 130342
            
            dResultadoValor = StrParaDbl(objResultadoTeste.sResultadoValor)
            
            If objResultadoTeste.iResultadoNaoConforme = 0 Then
                If dResultadoValor < (objResultadoTeste.dTesteLimiteDe - QTDE_ESTOQUE_DELTA) Or _
                        dResultadoValor > (objResultadoTeste.dTesteLimiteAte + QTDE_ESTOQUE_DELTA) Then gError 130339
            Else
                If dResultadoValor > (objResultadoTeste.dTesteLimiteDe - QTDE_ESTOQUE_DELTA) And _
                        dResultadoValor < (objResultadoTeste.dTesteLimiteAte + QTDE_ESTOQUE_DELTA) Then gError 130340
            End If
        End If

    End If
    
    Move_LinhaTeste_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_LinhaTeste_Memoria:

    Move_LinhaTeste_Memoria = gErr
     
    Select Case gErr
          
        Case 95462
        
        Case 95463
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTESQUALIDADE_NAO_CADASTRADO", gErr, objTeste.sNomeReduzido)
          
        Case 130339
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_QUALIDADE_NAO_CONFORME", gErr, objTeste.iCodigo, iLinha)
          
        Case 130340
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_QUALIDADE_CONFORME", gErr, objTeste.iCodigo, iLinha)
        
        Case 130341
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_QUALIDADE_SEM_RESULTADO", gErr, objTeste.iCodigo, iLinha)
        
        Case 130342
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_QUALIDADE_NAO_NUMERICO", gErr, objTeste.iCodigo, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165909)
     
    End Select
     
    Exit Function

End Function

Function Saida_Celula_Limite(objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Limite
    
    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then
            
        lErro = Valor_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 130314
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Limite = SUCESSO

    Exit Function

Erro_Saida_Celula_Limite:

    Saida_Celula_Limite = gErr

    Select Case gErr

        Case 95082, 130314
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165910)

    End Select

    Exit Function

End Function

Public Sub LimiteDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LimiteDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub LimiteDe_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub LimiteDe_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub LimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = LimiteDe
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub LimiteAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LimiteAte_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub LimiteAte_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub LimiteAte_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub LimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = LimiteAte
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub NoCertificado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NoCertificado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub NoCertificado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub NoCertificado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub NoCertificado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = NoCertificado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Especificacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Especificacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Especificacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub Especificacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub Especificacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Especificacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ObservacaoTeste_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ObservacaoTeste_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ObservacaoTeste_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub ObservacaoTeste_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub ObservacaoTeste_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ObservacaoTeste
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Metodo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Metodo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Metodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub Metodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub Metodo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Metodo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoTestes_Click()

Dim objTeste As New ClassTestesQualidade
Dim colSelecao As New Collection
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoTestes_Click

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95121

    'Verifica se há uma linha selecionada
    If GridTestes.Row = 0 Then gError 95108
        
    'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
    faz uma nova leitura em busca do codigo
    objTeste.sNomeReduzido = GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col)
    lErro = CF("TesteQualidade_Le_NomeReduzido", objTeste)
    If lErro <> SUCESSO And lErro <> 130109 Then gError 95468
    
    'chama a tela de browser
    Call Chama_Tela("TestesQualidadeLista", colSelecao, objTeste, objEventoTestes)
    
    iAlterado = 1
    
    Exit Sub
    
Erro_BotaoTestes_Click:
    
    Select Case gErr
    
        Case 95468
        
        Case 95469
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTeste.sNomeReduzido)
            
            If vbMsgRes = vbYes Then
                'Chama a tela de Testes
                Call Chama_Tela("TestesQualidade", objTeste)
            End If
        
        Case 95108
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 95121
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165911)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoTestes_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTeste As ClassTestesQualidade
Dim iLinha As Integer
Dim sSiglaAux As String
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoTestes_evSelecao
           
    'Verifica se o teste da linha atual está preenchido
    If Len(Trim(GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col))) = 0 Then
            
        'Define o tipo de obj recebido (Tipo Teste)
        Set objTeste = obj1
        
        'Verifica se há algumo teste repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridTestes.Row Then
                
                If GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido Then
                    Teste.Text = ""
                    gError 95120
                End If
                    
            End If
                           
        Next
        
        Teste.Text = objTeste.sNomeReduzido
            
        If Not (ActiveControl Is Teste) Then
        
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col) = objTeste.sNomeReduzido
        
            lErro = Teste_Saida_Celula(GridTestes.Row, objTeste)
            If lErro <> SUCESSO Then gError 130338
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTestes_evSelecao:

    Select Case gErr

        Case 95464, 130338
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objTeste.iCodigo)
        
        Case 95120
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_REPETIDA", gErr, objTeste.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGrid)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165912)
              
    End Select
    
    Exit Sub
        
End Sub

Public Sub Teste_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Teste_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Teste_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub Teste_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub Teste_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Teste
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ResultadoNaoConforme_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoNaoConforme_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoNaoConforme_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub ResultadoNaoConforme_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub ResultadoNaoConforme_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ResultadoNaoConforme
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ResultadoValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoValor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub ResultadoValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub ResultadoValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ResultadoValor
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub RegistroAnaliseID_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub RegistroAnaliseID_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub RegistroAnaliseID_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub RegistroAnaliseID_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub RegistroAnaliseID_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = RegistroAnaliseID
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub RegistroAnaliseData_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub RegistroAnaliseData_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub RegistroAnaliseData_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub RegistroAnaliseData_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub RegistroAnaliseData_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = RegistroAnaliseData
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Traz_Qualidade_Tela(objRastroLote As ClassRastreamentoLote) As Long

Dim lErro As Long, iLinha As Integer, objResultadoTeste As ClassRastreamentoLoteTeste
Dim objTeste As New ClassTestesQualidade

On Error GoTo Erro_Traz_Qualidade_Tela

    For Each objResultadoTeste In objRastroLote.colTeste
        
        'Incrementa o contador
        iLinha = iLinha + 1
        
        'Guarda em objTeste o código do teste que está em objProdutoTeste
        objTeste.iCodigo = objResultadoTeste.iTesteCodigo
        
        'Le do BD a descricao do teste
        lErro = CF("TestesQualidade_Le", objTeste)
        If lErro <> SUCESSO And lErro <> 130109 Then gError 95048
        
        'Se nao encontrou => erro
        If lErro <> SUCESSO Then gError 95049
        
        'Exibe o teste
        GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido
        
        'Exibir o restante
        GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoNaoConforme_Col) = objResultadoTeste.iResultadoNaoConforme
        GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoValor_Col) = objResultadoTeste.sResultadoValor
        GridTestes.TextMatrix(iLinha, iGridTestes_ResultadoObs_Col) = objResultadoTeste.sResultadoObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseID_Col) = objResultadoTeste.sRegistroAnaliseID
        If objResultadoTeste.dtRegistroAnaliseData <> DATA_NULA Then
            GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseData_Col) = Format(objResultadoTeste.dtRegistroAnaliseData, "dd/mm/yyyy")
        Else
            GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseData_Col) = ""
        End If
        GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col) = objResultadoTeste.sTesteEspecificacao
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col) = Format(objResultadoTeste.dTesteLimiteDe, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col) = Format(objResultadoTeste.dTesteLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col) = objResultadoTeste.sTesteMetodoUsado
        GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col) = objResultadoTeste.sTesteObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col) = objResultadoTeste.iTesteNoCertificado
    
    Next
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    objGrid.iLinhasExistentes = iLinha
    
    Traz_Qualidade_Tela = SUCESSO
     
    Exit Function
    
Erro_Traz_Qualidade_Tela:

    Traz_Qualidade_Tela = gErr
     
    Select Case gErr
          
        Case 130317, 95048
        
        Case 95049
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTESQUALIDADE_NAO_PREENCHIDO", gErr, objTeste.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165913)
     
    End Select
     
    Exit Function

End Function

Public Sub ResultadoObs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoObs_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub ResultadoObs_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Public Sub ResultadoObs_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Public Sub ResultadoObs_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ResultadoObs
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Teste_ExibeInfo(ByVal iLinha As Integer)

Dim lErro As Long

On Error GoTo Erro_Teste_ExibeInfo

    If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) > 0 Then
    
        LabelTeste.Caption = GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col)
        LabelLimiteDe = GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col)
        LabelLimiteAte = GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col)
        NoCertificado2.Value = IIf(StrParaInt(GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col)) <> 0, vbChecked, vbUnchecked)
        LabelMetodo = GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col)
        LabelEspecificacao = GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col)
        LabelObservacao = GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col)
    
    Else
        
        Call Teste_LimpaInfo
    
    End If
    
    Exit Sub
     
Erro_Teste_ExibeInfo:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165914)
     
    End Select
     
    Exit Sub

End Sub

Sub Teste_LimpaInfo()

    LabelTeste.Caption = ""
    LabelLimiteDe = ""
    LabelLimiteAte = ""
    NoCertificado2.Value = vbUnchecked
    LabelMetodo = ""
    LabelEspecificacao = ""
    LabelObservacao = ""
    
End Sub

Private Function Teste_Saida_Celula(ByVal iLinha As Integer, ByVal objTeste As ClassTestesQualidade) As Long
'completa com dados de produtoteste ou, se nao houver, com dados do proprio teste (??? falta parte)

Dim lErro As Long
Dim objProdutoTeste As New ClassProdutoTeste, sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Teste_Saida_Celula

    'Preenche o grid
    Teste.Text = objTeste.sNomeReduzido
        
    If Len(Trim(IdAnalisePadrao.Text)) > 0 Then
        GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseID_Col) = IdAnalisePadrao.Text
    End If

    If StrParaDate(Data.Text) <> DATA_NULA Then
        GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseData_Col) = Data.Text
    End If

    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 130335

    objProdutoTeste.sProduto = sProdutoFormatado
    objProdutoTeste.iTesteCodigo = objTeste.iCodigo
    
    lErro = CF("ProdutoTeste_Le", objProdutoTeste)
    If lErro <> SUCESSO And lErro <> 130292 Then gError 130336
    
    If lErro = SUCESSO Then
    
        'Exibir o restante
        GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col) = objProdutoTeste.sTesteEspecificacao
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col) = Format(objProdutoTeste.dTesteLimiteDe, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col) = Format(objProdutoTeste.dTesteLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col) = objProdutoTeste.sTesteMetodoUsado
        GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col) = objProdutoTeste.sTesteObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col) = objProdutoTeste.iTesteNoCertificado
        
    Else
    
        GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col) = objTeste.sEspecificacao
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col) = Format(objTeste.dLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col) = Format(objTeste.dLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col) = objTeste.sMetodoUsado
        GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col) = objTeste.sObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col) = objTeste.iNoCertificado
        
    End If
    
    'Se necessário cria uma nova linha no Grid
    If GridTestes.Row - GridTestes.FixedRows = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
    End If
    
    Call Grid_Refresh_Checkbox(objGrid)
        
    Teste_Saida_Celula = SUCESSO
     
    Exit Function
    
Erro_Teste_Saida_Celula:

    Teste_Saida_Celula = gErr
     
    Select Case gErr
          
        Case 130335, 130336
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165915)
     
    End Select
     
    Exit Function

End Function

'#############################################
'Inserido por Wagner
Public Sub BotaoTrazerTeste_Click()

Dim lErro As Long
Dim objProdutoTeste As ClassProdutoTeste
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim objTeste As ClassTestesQualidade

On Error GoTo Erro_BotaoTrazerTeste_Click

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 138296

    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 138297

    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("ProdutoTeste_Le_Produto", objProduto)
    If lErro <> SUCESSO And lErro <> 130294 Then gError 138298
    
    iLinha = 0
    
    Call Grid_Limpa(objGrid)
    
    For Each objProdutoTeste In objProduto.colProdutoTeste
    
        Set objTeste = New ClassTestesQualidade
    
        iLinha = iLinha + 1
    
        'Guarda em objTeste o código do teste que está em objProdutoTeste
        objTeste.iCodigo = objProdutoTeste.iTesteCodigo
    
        'Le do BD a descricao do teste
        lErro = CF("TestesQualidade_Le", objTeste)
        If lErro <> SUCESSO And lErro <> 130109 Then gError 138299
        
        'Se nao encontrou => erro
        If lErro <> SUCESSO Then gError 138300
        
        'Exibe o teste
        GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido
        
        'Exibir o restante
        GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col) = objProdutoTeste.sTesteEspecificacao
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col) = Format(objProdutoTeste.dTesteLimiteDe, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col) = Format(objProdutoTeste.dTesteLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col) = objProdutoTeste.sTesteMetodoUsado
        GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col) = objProdutoTeste.sTesteObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col) = objProdutoTeste.iTesteNoCertificado
        
        If Len(Trim(IdAnalisePadrao.Text)) > 0 Then
            GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseID_Col) = IdAnalisePadrao.Text
        End If
        
        If StrParaDate(Data.Text) <> DATA_NULA Then
            GridTestes.TextMatrix(iLinha, iGridTestes_RegistroAnaliseData_Col) = Data.Text
        End If
    
    Next

    Call Grid_Refresh_Checkbox(objGrid)
    
    objGrid.iLinhasExistentes = iLinha

    Exit Sub
     
Erro_BotaoTrazerTeste_Click:

    Select Case gErr
          
        Case 138296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
          
        Case 138297 To 138299
        
        Case 138300
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTESQUALIDADE_NAO_CADASTRADO", gErr, objTeste.iCodigo)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165916)
     
    End Select
     
    Exit Sub
    
End Sub
'#############################################

'#############################################
'Inserido por Wagner 29/09/05
Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 140410

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 140410

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165917)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 140411

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 140411

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165918)

    End Select

    Exit Sub

End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub IdAnalisePadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 140412

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 140412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165919)

    End Select

    Exit Sub

End Sub
'#############################################
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

Public Function RL_Move_Tela_Memoria(ByVal objCT As Object, objRastroLote As ClassRastreamentoLote) As Long
    RL_Move_Tela_Memoria = SUCESSO
End Function

Public Function RL_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    RL_Form_Unload = SUCESSO
End Function

Public Function RL_Traz_RastroLote_Tela(ByVal objCT As Object, objRastroLote As ClassRastreamentoLote) As Long
    RL_Traz_RastroLote_Tela = SUCESSO
End Function

Public Function RL_Limpa_Tela_RastroLote(ByVal objCT As Object) As Long
    RL_Limpa_Tela_RastroLote = SUCESSO
End Function

Public Function RL_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long
    RL_UserControl_KeyDown = SUCESSO
End Function

Public Sub Localizacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cliente_Change()
    iClienteAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.lCodigo
    
    'Dispara o Validate de Cliente
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o codigo do Cliente
    objCliente.lCodigo = Codigo_Extrai(Cliente.Text)

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate
   
    If iClienteAlterado <> 0 Then

        FilialCli.Clear
        
        'Se o Cliente foi preenchido
        If Len(Trim(Cliente.Text)) > 0 Then
    
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 6668 Then gError 116445
        
            Cliente.Text = objCliente.sNomeReduzido
        
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 116446

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialCli, colCodigoNome)
    
        End If
        
    End If

    iClienteAlterado = 0
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 116445, 116446
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159139)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialCli_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_FilialCli_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCli.Text)) = 0 Then Exit Sub

    'verifica se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 119522

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If FilialCli.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(FilialCli, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 116447

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objFilialCliente.iCodFilial = iCodigo

        'Tentativa de leitura da Filial com esse código no BD
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 116448

        If lErro = 12567 Then gError 119521  'Não encontrou Filial no  BD

        'Encontrou Filial no BD, coloca no Text da Combo
        FilialCli.Text = objFilialCliente.sNome

    End If
        
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 116449

    Exit Sub

Erro_FilialCli_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 116447, 116448

        Case 119522
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 116449, 119521
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialCli.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159140)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLocalizacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastreamentoLoteLoc As ClassRastreamentoLoteLoc

On Error GoTo Erro_objEventoLocalizacao_evSelecao

    Set objRastreamentoLoteLoc = obj1

    Localizacao.Text = objRastreamentoLoteLoc.sLocalizacao

    Me.Show

    Exit Sub

Erro_objEventoLocalizacao_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198730)

    End Select

    Exit Sub

End Sub

Public Sub LabelLocalizacao_Click()

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc
Dim colSelecao As New Collection

On Error GoTo Erro_LabelLocalizacao_Click

    'Verifica se o Localizacao foi preenchido
    If Len(Trim(Localizacao.Text)) <> 0 Then

        objRastreamentoLoteLoc.sLocalizacao = Localizacao.Text

    End If

    Call Chama_Tela("RastreamentoLoteLocLista", colSelecao, objRastreamentoLoteLoc, objEventoLocalizacao)

    Exit Sub

Erro_LabelLocalizacao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198731)

    End Select

    Exit Sub

End Sub

Public Sub Localizacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRastroLoteLoc As New ClassRastreamentoLoteLoc
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Localizacao_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(Trim(Localizacao.Text)) <> 0 Then
    
        objRastroLoteLoc.sLocalizacao = Localizacao.Text
    
        lErro = CF("RastreamentoLoteLoc_Le", objRastroLoteLoc)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198732
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 198733
    
    End If
    
    Exit Sub

Erro_Localizacao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 198732

        Case 198733
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_RASTROLOCALIZACAO", objRastroLoteLoc.sLocalizacao)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("RastreamentoLoteLoc", objRastroLoteLoc)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 198734)

    End Select

    Exit Sub

End Sub

