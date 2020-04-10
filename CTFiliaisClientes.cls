VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTFiliaisClientes"
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

'DECLARAÇÃO DE VARIÁVEIS GLOBAIS
Public gobjTabEnd As New ClassTabEndereco

Dim iFrameAtual1 As Integer
Dim iFrameAtual2 As Integer
Public iAlterado As Integer
Dim iVendedorAlterado As Integer

Dim sIEAnt As String

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCobrador As AdmEvento
Attribute objEventoCobrador.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFilialCliente As AdmEvento
Attribute objEventoFilialCliente.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento 'Inserido por Wagner 26/09/05
Attribute objEventoMensagem.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Inscricoes = 2
Private Const TAB_Enderecos = 3
Private Const TAB_Vendas = 4
Private Const TAB_Estatisticas = 5

'--- inicio dos properties get dos controles da tela
Public Property Get IEIsento() As Object
     Set IEIsento = objUserControl.Controls("IEIsento")
End Property

Public Property Get IENaoContrib() As Object
     Set IENaoContrib = objUserControl.Controls("IENaoContrib")
End Property

Public Property Get RedespachoCli() As Object
     Set RedespachoCli = objUserControl.Controls("RedespachoCli")
End Property

Public Property Get Guia() As Object
     Set Guia = objUserControl.Controls("Guia")
End Property

Public Property Get TranspRedespacho() As Object
     Set TranspRedespacho = objUserControl.Controls("TranspRedespacho")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get Ativo() As Object
     Set Ativo = objUserControl.Controls("Ativo")
End Property

Public Property Get ComboCategoriaClienteItem() As Object
     Set ComboCategoriaClienteItem = objUserControl.Controls("ComboCategoriaClienteItem")
End Property

Public Property Get ComboCategoriaCliente() As Object
     Set ComboCategoriaCliente = objUserControl.Controls("ComboCategoriaCliente")
End Property

Public Property Get GridCategoria() As Object
     Set GridCategoria = objUserControl.Controls("GridCategoria")
End Property

Public Property Get Filiais() As Object
     Set Filiais = objUserControl.Controls("Filiais")
End Property

Public Property Get CodFilial() As Object
     Set CodFilial = objUserControl.Controls("CodFilial")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get CodCliente() As Object
     Set CodCliente = objUserControl.Controls("CodCliente")
End Property

Public Property Get LabelCodCliente() As Object
     Set LabelCodCliente = objUserControl.Controls("LabelCodCliente")
End Property

Public Property Get LabelCodFilial() As Object
     Set LabelCodFilial = objUserControl.Controls("LabelCodFilial")
End Property

Public Property Get Nome() As Object
     Set Nome = objUserControl.Controls("Nome")
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get CGC() As Object
     Set CGC = objUserControl.Controls("CGC")
End Property

Public Property Get RG() As Object
     Set RG = objUserControl.Controls("RG")
End Property

Public Property Get InscricaoEstadual() As Object
     Set InscricaoEstadual = objUserControl.Controls("InscricaoEstadual")
End Property

Public Property Get InscricaoMunicipal() As Object
     Set InscricaoMunicipal = objUserControl.Controls("InscricaoMunicipal")
End Property

Public Property Get InscricaoSuframa() As Object
     Set InscricaoSuframa = objUserControl.Controls("InscricaoSuframa")
End Property

Public Property Get FilialLabel3() As Object
     Set FilialLabel3 = objUserControl.Controls("FilialLabel3")
End Property

Public Property Get FilialLabel() As Object
     Set FilialLabel = objUserControl.Controls("FilialLabel")
End Property

Public Property Get DataUltimaCompra() As Object
     Set DataUltimaCompra = objUserControl.Controls("DataUltimaCompra")
End Property

Public Property Get DataPrimeiraCompra() As Object
     Set DataPrimeiraCompra = objUserControl.Controls("DataPrimeiraCompra")
End Property

Public Property Get ValorAcumuladoCompras() As Object
     Set ValorAcumuladoCompras = objUserControl.Controls("ValorAcumuladoCompras")
End Property

Public Property Get MediaCompra() As Object
     Set MediaCompra = objUserControl.Controls("MediaCompra")
End Property

Public Property Get NumeroCompras() As Object
     Set NumeroCompras = objUserControl.Controls("NumeroCompras")
End Property

Public Property Get DataUltChequeDevolvido() As Object
     Set DataUltChequeDevolvido = objUserControl.Controls("DataUltChequeDevolvido")
End Property

Public Property Get NumChequesDevolvidos() As Object
     Set NumChequesDevolvidos = objUserControl.Controls("NumChequesDevolvidos")
End Property

Public Property Get MaiorAtraso() As Object
     Set MaiorAtraso = objUserControl.Controls("MaiorAtraso")
End Property

Public Property Get MediaAtraso() As Object
     Set MediaAtraso = objUserControl.Controls("MediaAtraso")
End Property

Public Property Get ValorPagtosAtraso() As Object
     Set ValorPagtosAtraso = objUserControl.Controls("ValorPagtosAtraso")
End Property

Public Property Get SaldoAtrasados() As Object
     Set SaldoAtrasados = objUserControl.Controls("SaldoAtrasados")
End Property

Public Property Get SaldoPedidosLiberados() As Object
     Set SaldoPedidosLiberados = objUserControl.Controls("SaldoPedidosLiberados")
End Property

Public Property Get SaldoDuplicatas() As Object
     Set SaldoDuplicatas = objUserControl.Controls("SaldoDuplicatas")
End Property

Public Property Get SaldoTitulos() As Object
     Set SaldoTitulos = objUserControl.Controls("SaldoTitulos")
End Property

Public Property Get Regiao() As Object
     Set Regiao = objUserControl.Controls("Regiao")
End Property

Public Property Get Cobrador() As Object
     Set Cobrador = objUserControl.Controls("Cobrador")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get ComissaoVendas() As Object
     Set ComissaoVendas = objUserControl.Controls("ComissaoVendas")
End Property

Public Property Get FreqVisitas() As Object
     Set FreqVisitas = objUserControl.Controls("FreqVisitas")
End Property

Public Property Get DataUltVisita() As Object
     Set DataUltVisita = objUserControl.Controls("DataUltVisita")
End Property

Public Property Get FilialLabel1() As Object
     Set FilialLabel1 = objUserControl.Controls("FilialLabel1")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get VendedorLabel() As Object
     Set VendedorLabel = objUserControl.Controls("VendedorLabel")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get LabelCobrador() As Object
     Set LabelCobrador = objUserControl.Controls("LabelCobrador")
End Property

Public Property Get OpcaoEndereco() As Object
     Set OpcaoEndereco = objUserControl.Controls("OpcaoEndereco")
End Property

Public Property Get FilialLabel2() As Object
     Set FilialLabel2 = objUserControl.Controls("FilialLabel2")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
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

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get MensagemNFLabel() As Object
     Set MensagemNFLabel = objUserControl.Controls("MensagemNFLabel")
End Property

Public Property Get TabEnd() As Object
     Set TabEnd = objUserControl.Controls("TabEnd")
End Property

Public Property Get RegimeTributario() As Object
     Set RegimeTributario = objUserControl.Controls("RegimeTributario")
End Property


'--- fim dos properties get dos controles da tela

Public Sub LabelCodCliente_Click()

Dim colSelecao As Collection
Dim objCliente As New ClassCliente

    'Preenche ClienteAte com o cliente da tela
    objCliente.lCodigo = StrParaLong(CodCliente.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Agente Cobrador
    CodCliente.Text = CStr(objCliente.lCodigo)

    Call CodCliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCodFilial_Click()

Dim colSelecao As Collection
Dim objFilialCliente As New ClassFilialCliente

    'Preenche ClienteAte com o cliente da tela
    objFilialCliente.lCodCliente = StrParaLong(CodCliente.Text)
    objFilialCliente.iCodFilial = StrParaInt(CodFilial.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("FiliaisClientesLista", colSelecao, objFilialCliente, objEventoFilialCliente)

End Sub

Private Sub objEventoFilialCliente_evSelecao(obj1 As Object)

Dim objFilialCliente As ClassFilialCliente
Dim objFilialClienteEstatistica As New ClassFilialClienteEst
Dim lErro As Long

On Error GoTo Erro_objEventoFilialCliente_evSelecao

    Set objFilialCliente = obj1

    'Tenta ler Filial de Cliente com a chave passada em objFilialCliente
    lErro = CF("FilialCliente_Le_Estendida", objFilialCliente, objFilialClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52591 Then gError 83088

    'Se Filial não existe
    If lErro <> SUCESSO Then gError 83089

    'Exibe dados da Filial na Tela
    lErro = Exibe_Dados_FilialCliente(objFilialCliente, objFilialClienteEstatistica)
    If lErro <> SUCESSO Then gError 83090

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoFilialCliente_evSelecao:

    Select Case gErr

        Case 83088, 83090

        Case 83089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155814)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático de Filial
    lErro = CF("FiliaisClientes_Automatico", CLng(CodCliente.Text), iCodigo)
    If lErro <> SUCESSO Then Error 57530

    'Coloca código gerado na Tela
    CodFilial.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57530
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155815)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(CodCliente.Text)) = 0 Then Error 12655

    'Verifica se o código da Filial foi preenchido
    If Len(Trim(CodFilial.Text)) = 0 Then Error 12656

    'Verifica se é Matriz
    If CInt(CodFilial.Text) = FILIAL_MATRIZ Then Error 12715

    objCliente.lCodigo = CLng(CodCliente.Text)
    objFilialCliente.lCodCliente = CLng(CodCliente.Text)
    objFilialCliente.iCodFilial = CInt(CodFilial.Text)

    'Lê os dados do Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 12657

    'Verifica se Cliente não está cadastrado
    If lErro <> SUCESSO Then Error 12658

    'Lê os dados da Filial Cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then Error 12659

    'Verifica se a Filial Cliente não está cadastrada
    If lErro <> SUCESSO Then Error 12660

    'Envia aviso perguntando se realmente deseja excluir Filial Cliente
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FILIALCLIENTE")

    If vbMsgRes = vbYes Then

        'Exclui Filial de Cliente
        lErro = CF("FilialCliente_Exclui", objFilialCliente)
        If lErro <> SUCESSO Then Error 12669

'        'Exclui Filial da TreeView
'        Call Arvore_Excluir(Filiais, objFilialCliente)
'
        'Limpa a Tela
        lErro = Limpa_Tela_FiliaisClientes()
        If lErro <> SUCESSO Then Error 58593
        
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 12655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", Err)

        Case 12656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", Err)

        Case 12657, 12659, 12669, 58593

        Case 12658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case 12660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_EXCLUSAO_MATRIZ", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155816)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama função de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 12610

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisClientes()
    If lErro <> SUCESSO Then Error 58594
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 12610, 58594

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155817)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar modificações feitas
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 12710

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisClientes()
    If lErro <> SUCESSO Then Error 58595

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 12710 'cancelou operacao de gravacao , continua execucao normal
        
        Case 58595 'Tratado na Rotina Chamada
        
        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155818)

    End Select

End Sub

Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

    Case STRING_CPF 'CPF

        'Critica CPF
        lErro = Cpf_Critica(CGC.Text)
        If lErro <> SUCESSO Then Error 12586

        CGC.Format = "000\.000\.000-00; ; ; "
        CGC.Text = CGC.Text

    Case STRING_CGC 'CGC

        'Critica CGC
        lErro = Cgc_Critica(CGC.Text)
        If lErro <> SUCESSO Then Error 12587

        CGC.Format = "00\.000\.000\/0000-00; ; ; "
        CGC.Text = CGC.Text

    Case Else

        Error 12588

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case Err

        Case 12586, 12587

        Case 12588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155819)

    End Select


    Exit Sub

End Sub

Public Sub RG_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RG_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)

End Sub

Public Sub Cidade_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCobrador As New ClassCobrador
Dim iCodigo As Integer

On Error GoTo Erro_Cobrador_Validate

    'Verifica se foi preenchida a ComboBox Cobrador
    If Len(Trim(Cobrador.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Cobrador
    If Cobrador.Text = Cobrador.List(Cobrador.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Cobrador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19255

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        'Lê Cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 19256

        If lErro <> SUCESSO Then Error 19257 'Não encontrou Cobrador no BD

        'Encontrou Cobrador no BD, coloca no Text da Combo
        Cobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19258

    Exit Sub

Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19255, 19256

    Case 19257  'Não encontrou Cobrador no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Cobradores", objCobrador)

        End If

    Case 19258

        lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, Cobrador.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155820)

    End Select

    Exit Sub

End Sub

Public Sub CodCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodCliente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)

End Sub

Public Sub CodCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente
Dim iIndice As Integer

On Error GoTo Erro_CodCliente_Validate

    'Verifica se foi preenchido o campo CodCliente
    If Len(Trim(CodCliente.Text)) = 0 Then Exit Sub

    'Critica se é do tipo Long positivo
    lErro = Long_Critica(CodCliente.Text)
    If lErro <> SUCESSO Then Error 19101

    objCliente.lCodigo = CLng(CodCliente.Text)

    'Lê o Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 12571

    'Coloca o Nome Reduzido nas labels (mesmo se estiverem vazios)
    For iIndice = 0 To 3
        ClienteLabel(iIndice).Caption = objCliente.sNomeReduzido
    Next

    'Verifica se existe
    If lErro <> SUCESSO Then

        'Envia aviso perguntando se deseja cadastrar novo Cliente
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE")

        If vbMsgRes = vbYes Then

            'Chama tela de Clientes
            lErro = Chama_Tela("Clientes", objCliente)
            If lErro <> SUCESSO Then Error 12572

        Else

            Cancel = True

        End If

    End If

    Exit Sub

Erro_CodCliente_Validate:

    Cancel = True

    Select Case Err

        Case 12571, 12572, 19101

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155821)

    End Select

    Exit Sub

End Sub

Public Sub CodFilial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodFilial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFilial, iAlterado)

End Sub

Public Sub CodFilial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFilial_Validate

    'Verifica se foi preenchido o campo Codigo Filial
    If Len(Trim(CodFilial.Text)) = 0 Then Exit Sub

    'Verifica se é do tipo Inteiro e positivo
    lErro = Inteiro_Critica(CodFilial.Text)
    If lErro <> SUCESSO Then Error 19303

    Exit Sub

Erro_CodFilial_Validate:

    Cancel = True


    Select Case Err

        Case 19303

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155822)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaCliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Public Sub ComboCategoriaCliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaCliente
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComboCategoriaClienteItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaClienteItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaClienteItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaClienteItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    lErro = Trata_ComboCategoriaClienteItem()
    If lErro <> SUCESSO Then Error 33079

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaClienteItem_GotFocus:

    Select Case Err

        Case 33079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155823)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaClienteItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaClienteItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaClienteItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Trata_ComboCategoriaClienteItem()

Dim lErro As Long
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim iIndice As Integer, sValor As String

On Error GoTo Erro_Trata_ComboCategoriaClienteItem

    'Armazena o Item em sValor
    sValor = ComboCategoriaClienteItem.Text
    
    'Limpa a Combo
    ComboCategoriaClienteItem.Clear

    'Retorna o Item para a Combo
    ComboCategoriaClienteItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(Trim(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col))) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaCliente.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaClienteItem(objCategoriaCliente)
        If lErro <> SUCESSO Then Error 33080

    End If

    'Seleciona Item que está no Grid
    For iIndice = 0 To ComboCategoriaClienteItem.ListCount - 1
        If ComboCategoriaClienteItem.List(iIndice) = GridCategoria.Text Then
            ComboCategoriaClienteItem.ListIndex = iIndice
            Exit For
        End If
    Next

    Trata_ComboCategoriaClienteItem = SUCESSO

    Exit Function

Erro_Trata_ComboCategoriaClienteItem:

    Trata_ComboCategoriaClienteItem = Err

    Select Case Err

        Case 33080

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155824)

    End Select

End Function

Private Function Carrega_ComboCategoriaClienteItem(objCategoriaCliente As ClassCategoriaCliente) As Long
'Carrega os Itens da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaClienteItem As ClassCategoriaClienteItem

On Error GoTo Erro_Carrega_ComboCategoriaClienteItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then Error 33081

    'Insere na combo CategoriaClienteItem
    For Each objCategoriaClienteItem In colItensCategoria

        'Insere na combo CategoriaCliente
        ComboCategoriaClienteItem.AddItem objCategoriaClienteItem.sItem

    Next

    Carrega_ComboCategoriaClienteItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaClienteItem:

    Carrega_ComboCategoriaClienteItem = Err

    Select Case Err

        Case 33081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155825)

    End Select

    Exit Function

End Function

Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39811

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39812

        'Carrega a conta formatada na tela
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica a conta
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 12596

        If lErro = 5700 Then Error 12597

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case Err

        Case 12596, 39811
    
        Case 12597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
    
        Case 39812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155826)
    
    End Select

    Exit Sub

End Sub

Public Sub ComissaoVendas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoVendas_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sComissaoVendas As String

On Error GoTo Erro_ComissaoVendas_Validate

    sComissaoVendas = ComissaoVendas.Text

    'Verifica se foi preenchida a Comissão de Venda
    If Len(Trim(ComissaoVendas.Text)) = 0 Then Exit Sub

    'Critica porcentagem
    lErro = Porcentagem_Critica(ComissaoVendas.Text)
    If lErro <> SUCESSO Then Error 12598

    ComissaoVendas.Text = Format(sComissaoVendas, "Fixed")

    Exit Sub

Erro_ComissaoVendas_Validate:

    Cancel = True


    Select Case Err

        Case 12598
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155827)
    
    End Select

    Exit Sub

End Sub

Public Sub DataUltVisita_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataUltVisita_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataUltVisita, iAlterado)

End Sub

Public Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataUltVisita_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataUltVisita_Validate

    'Verifica se foi preenchido o Desconto
    If Len(Trim(DataUltVisita.Text)) = 0 Then Exit Sub

    'Critica se é do tipo Data
    lErro = Data_Critica(DataUltVisita.Text)
    If lErro <> SUCESSO Then Error 12605

    Exit Sub

Erro_DataUltVisita_Validate:

    Cancel = True


    Select Case Err

        Case 12605
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155828)
    
    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub FreqVisitas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FreqVisitas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FreqVisitas, iAlterado)

End Sub

Public Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Public Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Public Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub

Public Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Public Sub GridCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Public Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Public Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Public Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As Collection

    'Se a transportadora estiver preenchida recolhe.
    If Len(Trim(Transportadora.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    'Chama a Tela de browse de Transportadora
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

    Exit Sub

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportdora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCobrador_Click()

Dim objCobrador As New ClassCobrador
Dim colSelecao As New Collection
    
    'Se o cobrador estiver preenchida passa o codigo para objCobrador
    If Len(Trim(Cobrador.Text)) > 0 Then objCobrador.iCodigo = Codigo_Extrai(Cobrador.Text)
    
    'Abre uma lista de Cobradores
    Call Chama_Tela("CobradorLista", colSelecao, objCobrador, objEventoCobrador)

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Regiao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Regiao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'Verifica se foi preenchido o campo Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Regiao
    If Regiao.Text = Regiao.List(Regiao.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19251

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 19252

        If lErro <> SUCESSO Then Error 19253 'Não encontrou Regiao Venda BD

        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19254

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19251, 19252

        Case 19253  'Não encontrou RegiaoVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

            If vbMsgRes = vbYes Then
                'Chama a tela de RegiaoVendasDeClientes
                Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
    
            End If
    
        Case 19254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", Err, Regiao.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155831)

    End Select

    Exit Sub

End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19259

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 19260
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then Error 19261

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19262

    Exit Sub

Erro_Transportadora_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19259, 19260
    
        Case 19261  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")
    
            If vbMsgRes = vbYes Then
    
                Call Chama_Tela("Transportadora", objTransportadora)
    
            End If
    
        Case 19262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155832)

    End Select

    Exit Sub

End Sub

Public Sub Vendedor_Change()

    iVendedorAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim dPercComissao As Double

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If iVendedorAlterado = 1 Then

        If Len(Trim(Vendedor.Text)) > 0 Then

            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Vendedor_Le(Vendedor, objVendedor)
            If lErro <> SUCESSO Then gError 25036

            lErro = CF("Vendedor_Comissao_Customiza", objVendedor, dPercComissao)
            If lErro <> SUCESSO Then gError 126954
            
            If Len(Trim(ComissaoVendas.Text)) = 0 And dPercComissao <> 0 Then ComissaoVendas.Text = CStr(100 * dPercComissao)

            If objVendedor.iAtivo = DESMARCADO Then gError 140221 'Inserido por Wagner

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 25036, 126954   'Tratado na rotina chamada
                        
        Case 140221
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155833)

    End Select

End Sub

Public Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub OpcaoEndereco_Click(Index As Integer)

    'Se frame selecionado for igual ao atual então não faz nada .
    If iFrameAtual2 <> Index Then
        'dependendo mostra o Endereço com o indice Selecionado e esconde os outros dois
        Select Case Index
            Case 0 'Principal
                Frame2(0).Visible = True
                Frame2(1).Visible = False
                Frame2(2).Visible = False
            Case 1 'Entrega
                Frame2(1).Visible = True
                Frame2(0).Visible = False
                Frame2(2).Visible = False
            Case 2 'Cobrança
                Frame2(2).Visible = True
                Frame2(0).Visible = False
                Frame2(1).Visible = False
        End Select
        'Armazena novo valor de iFrameAtual
        iFrameAtual2 = Index
    End If

End Sub

Public Sub Filiais_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim sCodFilial As String
Dim sCodCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Filiais_NodeClick

    'Se não deu click em Filial sai da rotina
    If Node.Parent Is Nothing Then Exit Sub

    'Determina chave CodCliente-CodFilial associada ao Nó
    sCodFilial = right(Node.Key, Len(Node.Key) - Len(Node.Parent.Key) - 1)
    sCodCliente = right(Node.Parent.Key, Len(Node.Parent.Key) - 1)

    objFilialCliente.lCodCliente = CLng(sCodCliente)
    objFilialCliente.iCodFilial = CInt(sCodFilial)

    'Tenta ler Filial de Cliente com a chave passada em objFilialCliente
    lErro = CF("FilialCliente_Le_Estendida", objFilialCliente, objFilialClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52591 Then Error 12568

    'Se Filial não existe
    If lErro <> SUCESSO Then Error 12569

    'Exibe dados da Filial na Tela
    lErro = Exibe_Dados_FilialCliente(objFilialCliente, objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 12570

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Filiais_NodeClick:

    Select Case Err

        Case 12568, 12570

        Case 12569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155834)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_FiliaisClientes() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sCodigoCliente As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FiliaisClientes

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Guarda o código de Cliente
    sCodigoCliente = CodCliente.Text

    'Limpa os MaskedEdit e TextBoxes
    Call Limpa_Tela(Me)
    
    Ativo.Value = MARCADO
    
    'Mantém o código de Cliente na Tela
    CodCliente.Text = sCodigoCliente
    
    'Limpa o código da Filial
    CodFilial.Text = ""

    Regiao.Text = ""
    Cobrador.Text = ""
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False

    'Limpa os Labels
    FilialLabel(0).Caption = ""
    FilialLabel(1).Caption = ""
    FilialLabel(2).Caption = ""

    'Escolhe opção Endereço Principal no OptionButton
    OpcaoEndereco(0) = True
    OpcaoEndereco(1) = False
    OpcaoEndereco(2) = False

    'Torna visível o Frame do Endereço Principal
    Frame2(0).Visible = True
    iFrameAtual2 = 0
    
    IEIsento.Value = vbChecked
    IENaoContrib.Value = vbChecked

    'Limpa o grid de Categorias
    Call Grid_Limpa(objGridCategoria)
    
    'limpa o tab de estatíticas:
    FilialLabel(3).Caption = ""
    SaldoTitulos.Caption = ""
    SaldoDuplicatas.Caption = ""
    SaldoPedidosLiberados.Caption = ""
    NumeroCompras.Caption = ""
    MediaCompra.Caption = ""
    ValorAcumuladoCompras.Caption = ""
    DataPrimeiraCompra.Caption = "  /  /  "
    DataUltimaCompra.Caption = "  /  /  "
    MediaAtraso.Caption = ""
    MaiorAtraso.Caption = ""
    SaldoAtrasados.Caption = ""
    ValorPagtosAtraso.Caption = ""
    NumChequesDevolvidos.Caption = ""
    DataUltChequeDevolvido.Caption = "  /  /  "
    TipoFrete.ListIndex = -1
        
    Mensagem.Text = "" 'Inserido por Wagner 26/09/05
    
    Call gobjTabEnd.Limpa_Tela
        
    Limpa_Tela_FiliaisClientes = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_FiliaisClientes:
    
    Limpa_Tela_FiliaisClientes = Err
    
    Select Case Err
        
        Case 58590
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155835)

    End Select
    
    Exit Function
        
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim sMascaraConta As String
Dim objCobrador  As ClassCobrador
Dim ColCobrador As New Collection
Dim objTela As Object

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
    Set objEventoCobrador = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoTranspRedespacho = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoFilialCliente = New AdmEvento
    Set objEventoMensagem = New AdmEvento 'Inserido por Wagner 26/09/05
    
    iFrameAtual1 = 1
    iFrameAtual2 = 0

    OpcaoEndereco(0) = True
    OpcaoEndereco(1) = False
    OpcaoEndereco(2) = False

    Ativo.Value = MARCADO
    
    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then gError 12542

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12544
    
        ContaContabil.Mask = sMascaraConta

    Else
       
       'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 12544
    
        ContaContabil.Mask = sMascaraConta
        
        'Desabilita conta contábil
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
    End If

    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then gError 12545

    'Preenche a ComboBox Regiao com os objetos da colecao colCodigoNome
    For Each objCodigoNome In colCodigoNome
        Regiao.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoNome.iCodigo
    Next

    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 12546

    'Preenche a ComboBox Cobrador com os objetos da colecao colCodigoNome
    For Each objCobrador In ColCobrador
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If
    Next

    'Lê códigos e nomes reduzidos da tabela Transportadoras
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 12547

    'Preenche a ComboBox Transportadora com os objetos da colecao colCodigoNome
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
        TranspRedespacho.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        TranspRedespacho.ItemData(TranspRedespacho.NewIndex) = objCodigoNome.iCodigo
    Next

    Set objGridCategoria = New AdmGrid

    'Inicialiaza o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 12365

    'Carrega a combobox de Categoria Produto
    lErro = Carrega_ComboCategoriaCliente()
    If lErro <> SUCESSO Then gError 12366
    
    'Preenche Combo Mensagem
    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela Mensagem
    lErro = CF("Cod_Nomes_Le", "Mensagens", "Codigo", "Descricao", STRING_NFISCAL_MENSAGEM, colCodigoNome)
    If lErro <> SUCESSO Then gError 12266

    'Preenche a ComboBox Mensagem com os objetos da colecao colCodigoDescricao
    For Each objCodigoNome In colCodigoNome
        Mensagem.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Mensagem.ItemData(Mensagem.NewIndex) = objCodigoNome.iCodigo
    Next
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207450
    
    'Prepara as Combos  do Tab de Inscricao
    lErro = Inicializa_Tab_Inscricao()
    If lErro <> SUCESSO Then gError 207401
    
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 12365, 12366, 12541, 12542, 12543, 12544, 12545, 12546, 12547, 25665, 12266, 207401, 207450

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155836)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFilialCliente As ClassFilialCliente) As Long

Dim lErro As Long
Dim objFilialClienteEstatistica As New ClassFilialClienteEst
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objFilialCliente Is Nothing) Then

        'Se foi passado código de Cliente
        If objFilialCliente.lCodCliente <> 0 Then

            'Se foi passado código de Filial
            If objFilialCliente.iCodFilial <> 0 Then

                'Tenta ler FilialCliente
                lErro = CF("FilialCliente_Le_Estendida", objFilialCliente, objFilialClienteEstatistica)
                If lErro <> SUCESSO And lErro <> 52591 Then Error 12548

                'Se a Filial não existir
                If lErro <> SUCESSO Then
                    
                    'Limpa a Tela
                    lErro = Limpa_Tela_FiliaisClientes()
                    If lErro <> SUCESSO Then Error 58591
                    
                    CodCliente.Text = CStr(objFilialCliente.lCodCliente)
                    CodFilial.Text = CStr(objFilialCliente.iCodFilial)
                    
                    
                Else  'Filial existe, então exibe seus dados

                    lErro = Exibe_Dados_FilialCliente(objFilialCliente, objFilialClienteEstatistica)
                    If lErro <> SUCESSO Then Error 12549

                End If

            Else 'Apenas o código do Cliente foi passado
                
                'Limpa Tela
                lErro = Limpa_Tela_FiliaisClientes()
                If lErro <> SUCESSO Then Error 58592
                
                CodCliente.Text = CStr(objFilialCliente.lCodCliente)

                'Inserido por Leo em 10/01/02
                If Len(Trim(objFilialCliente.sNomeReduzidoCli)) > 0 Then
                    
                    For iIndice = 0 To 4
                        ClienteLabel(iIndice).Caption = objFilialCliente.sNomeReduzidoCli
                    Next

                End If
                
                If Len(Trim(objFilialCliente.sNome)) > 0 Then
                    Nome.Text = objFilialCliente.sNome
                End If
                'Leo até aqui

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 12548, 12549, 58591, 58592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155837)

    End Select

    iAlterado = 0

    Exit Function

End Function
                
Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categorias

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaCliente.Name)
    objGridInt.colCampo.Add (ComboCategoriaClienteItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Function Carrega_ComboCategoriaCliente() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos", colCategorias)
    If lErro <> SUCESSO Then Error 33078

    For Each objCategoriaCliente In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next

    Carrega_ComboCategoriaCliente = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaCliente:

    Carrega_ComboCategoriaCliente = Err

    Select Case Err

        Case 33078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155838)

    End Select

    Exit Function

End Function

Public Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub InscricaoMunicipal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub InscricaoSuframa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Nome_Validate(Cancel As Boolean)

    FilialLabel(0).Caption = Trim(Nome.Text)
    FilialLabel(1).Caption = Trim(Nome.Text)
    FilialLabel(2).Caption = Trim(Nome.Text)
    FilialLabel(3).Caption = Trim(Nome.Text)
    
End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual1 Then

        If TabStrip_PodeTrocarTab(iFrameAtual1, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual1).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual1 = Opcao.SelectedItem.Index
            
        Select Case iFrameAtual1
            
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ID
                
            Case TAB_Inscricoes
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_INSCRICOES
                        
            Case TAB_Enderecos
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ENDERECOS
            
            Case TAB_Vendas
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_VENDAS
                
            Case TAB_Estatisticas
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ESTATISTICAS
        
        End Select
    
    End If

End Sub

Private Function Arvore_Inicializa(Filiais As TreeView) As Long
'Move os dados das Filiais de Clientes do Banco de Dados para a TreeView Filiais

Dim colClientes As New AdmCollCodigoNome
Dim colFiliais As New ColChaveFilial
Dim objCliente As AdmlCodigoNome
Dim objFilialCliente As ClassChaveFilial
Dim sKeyCliente As String
Dim sKeyFilial As String
Dim sPaiKey As String
Dim objNodePai As Node
Dim lErro As Long

On Error GoTo Erro_Arvore_Inicializa

    'Lê Códigos, NomesReduzidos de Clientes no BD
    lErro = CF("LCod_Nomes_Le", "Clientes", "Codigo", "NomeReduzido", STRING_CLIENTE_NOME_REDUZIDO, colClientes)
    If lErro <> SUCESSO Then Error 12539

    'Preenche todos os clientes na TreeView com seu Nome Reduzido
    For Each objCliente In colClientes

        'Chave do Cliente
        sKeyCliente = KEY_CARACTER & CStr(objCliente.lCodigo)

        'Adiciona o Nó na árvore
        Call Filiais.Nodes.Add(, , sKeyCliente, objCliente.sNome)

    Next

    'Leitura das Filiais no BD
    lErro = CF("FiliaisClientes_Le_Chave", colFiliais)
    If lErro <> SUCESSO Then Error 12540

    'Preenche todas as Filiais Clientes na TreeView
    For Each objFilialCliente In colFiliais

        'Chaves da Filial e do Cliente Pai
        sKeyFilial = KEY_CARACTER & CStr(objFilialCliente.lCodPai) & SEPARADOR & CStr(objFilialCliente.iCodFilial)
        sPaiKey = KEY_CARACTER & CStr(objFilialCliente.lCodPai)

        'Adiciona o Nó na árvore
        Call Filiais.Nodes.Add(sPaiKey, tvwChild, sKeyFilial, objFilialCliente.sNome)

    Next

    Arvore_Inicializa = SUCESSO

    Exit Function

Erro_Arvore_Inicializa:

    Arvore_Inicializa = Err

    Select Case Err

        Case 12539, 12540

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155839)

    End Select

    Exit Function

End Function

Private Sub Arvore_Ordena(Filiais As TreeView)

Dim objNode As Node

    Filiais.Sorted = True

    For Each objNode In Filiais.Nodes

       objNode.Sorted = True

    Next

End Sub

Function Exibe_Dados_FilialCliente(objFilialCliente As ClassFilialCliente, objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Exibe os dados da filial do cliente na tela

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco
Dim objCodigoNome As New AdmlCodigoNome
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim colCategoriaItem As New Collection
Dim bCancel As Boolean
Dim objCliente As New ClassCliente

On Error GoTo Erro_Exibe_Dados_FilialCliente

    objCliente.lEndereco = objFilialCliente.lEndereco
    objCliente.lEnderecoCobranca = objFilialCliente.lEnderecoCobranca
    objCliente.lEnderecoEntrega = objFilialCliente.lEnderecoEntrega

    objCodigoNome.lCodigo = objFilialCliente.lCodCliente

    'Lê Cliente a partir do Nome Reduzido
    lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
    If lErro <> SUCESSO And lErro <> 12553 Then Error 12557

    'Cliente não cadastrado
    If lErro = 12553 Then Error 6926

    For iIndice = 0 To 4
        ClienteLabel(iIndice).Caption = objCodigoNome.sNome
    Next

    'IDENTIFICACAO :

    CodCliente.Text = CStr(objFilialCliente.lCodCliente)
    CodFilial.Text = CStr(objFilialCliente.iCodFilial)
    Nome.Text = objFilialCliente.sNome
        
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then Error 12429
        
    'Lê as Categoria da Filial do Cliente
    lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
    If lErro <> SUCESSO Then Error 12429
    
    Ativo.Value = objFilialCliente.iAtivo
    'Limpa o Grid de Categorias
    Call Grid_Limpa(objGridCategoria)

    'Carrega o Grid Categoria
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = colCategoriaItem.Item(iIndice).sCategoria
            GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = colCategoriaItem.Item(iIndice).sItem
        Next
        objGridCategoria.iLinhasExistentes = colCategoriaItem.Count
    End If

    'INSCRIÇÕES

    FilialLabel(2).Caption = objFilialCliente.sNome
    
    RG.Text = objFilialCliente.sRG
    CGC.Text = objFilialCliente.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objFilialCliente.sInscricaoEstadual
    Call Trata_IE
    If objFilialCliente.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    If objFilialCliente.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If
    InscricaoMunicipal.Text = objFilialCliente.sInscricaoMunicipal
    Observacao.Text = objFilialCliente.sObservacao
    
    InscricaoSuframa.PromptInclude = False
    InscricaoSuframa.Text = objFilialCliente.sInscricaoSuframa
    InscricaoSuframa.PromptInclude = True
    
        'Busca no ItemData da Combo de Regime Tributario o código
    For iIndice = 0 To RegimeTributario.ListCount - 1
        If RegimeTributario.ItemData(iIndice) = objFilialCliente.iRegimeTributario Then
            RegimeTributario.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'ENDERECOS :

    FilialLabel(0).Caption = objFilialCliente.sNome

    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
    If lErro <> SUCESSO Then Error 12558

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 12558

    'VENDAS :

    FilialLabel(1).Caption = objFilialCliente.sNome

    If objFilialCliente.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objFilialCliente.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If

    ContaContabil.PromptInclude = False

    If Len(Trim(objFilialCliente.sContaContabil)) = 0 Then
        ContaContabil.Text = objFilialCliente.sContaContabil
    Else
        lErro = Mascara_RetornaContaEnxuta(objFilialCliente.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12560
        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    If objFilialCliente.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objFilialCliente.dComissaoVendas)
    End If

    If objFilialCliente.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objFilialCliente.iRegiao)
        Call Regiao_Validate(bCancel)
    End If

    If objFilialCliente.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objFilialCliente.iCodCobrador)
        Call Cobrador_Validate(bCancel)
    End If

    If objFilialCliente.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objFilialCliente.iFreqVisitas)
    End If

    DataUltVisita.PromptInclude = False
    If (objFilialCliente.dtDataUltVisita) = DATA_NULA Then
        DataUltVisita.Text = "  /  /  "
    Else
        DataUltVisita.Text = Format(objFilialCliente.dtDataUltVisita, "dd/mm/yy")
    End If

    DataUltVisita.PromptInclude = True
    
    'Tipo de Frete
    If objFilialCliente.iTipoFrete <> -1 Then
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objFilialCliente.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoFrete.ListIndex = -1
    End If
    
    If objFilialCliente.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objFilialCliente.iCodTransportadora)
        Call Transportadora_Validate(bCancel)
    End If
    
    Guia.Text = objFilialCliente.sGuia
    
    If objFilialCliente.iCodTranspRedesp = 0 Then
        TranspRedespacho.Text = ""
    Else
        TranspRedespacho.Text = CStr(objFilialCliente.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If
    
    RedespachoCli.Value = objFilialCliente.iDetPagFrete

    'ESTATISTICAS :
    
    FilialLabel(3).Caption = objFilialCliente.sNome

    SaldoTitulos.Caption = Format(objFilialClienteEstatistica.dSaldoTitulos, "Standard")
    SaldoDuplicatas.Caption = Format(objFilialClienteEstatistica.dSaldoDuplicatas, "Standard")
    SaldoPedidosLiberados.Caption = Format(objFilialClienteEstatistica.dSaldoPedidosLiberados, "Standard")
    NumeroCompras.Caption = CStr(objFilialClienteEstatistica.lNumeroCompras)
    MediaCompra.Caption = Format(objFilialClienteEstatistica.dMediaCompra, "Standard")
    ValorAcumuladoCompras.Caption = Format(objFilialClienteEstatistica.dValorAcumuladoCompras, "Standard")

    If objFilialClienteEstatistica.dtDataPrimeiraCompra = DATA_NULA Then
        DataPrimeiraCompra.Caption = "  /  /  "
    Else
        DataPrimeiraCompra.Caption = Format(objFilialClienteEstatistica.dtDataPrimeiraCompra, "dd/mm/yyyy")
    End If

    If objFilialClienteEstatistica.dtDataUltimaCompra = DATA_NULA Then
        DataUltimaCompra.Caption = "  /  /  "
    Else
        DataUltimaCompra.Caption = Format(objFilialClienteEstatistica.dtDataUltimaCompra, "dd/mm/yyyy")
    End If
    
    If objFilialClienteEstatistica.lMediaAtraso <> 1 Then
        MediaAtraso.Caption = CStr(objFilialClienteEstatistica.lMediaAtraso) & " dias"
    Else
        MediaAtraso.Caption = CStr(objFilialClienteEstatistica.lMediaAtraso) & " dia"
    End If
    
    If objFilialClienteEstatistica.lMaiorAtraso <> 1 Then
        MaiorAtraso.Caption = CStr(objFilialClienteEstatistica.lMaiorAtraso) & " dias"
    Else
        MaiorAtraso.Caption = CStr(objFilialClienteEstatistica.lMaiorAtraso) & " dia"
    End If
            
    SaldoAtrasados.Caption = Format(objFilialClienteEstatistica.dSaldoAtrasados, "Standard")
    ValorPagtosAtraso.Caption = Format(objFilialClienteEstatistica.dValPagtosAtraso, "Standard")

    NumChequesDevolvidos.Caption = CStr(objFilialClienteEstatistica.iNumChequesDevolvidos)

    If objFilialClienteEstatistica.dtDataUltChequeDevolvido = DATA_NULA Then
        DataUltChequeDevolvido.Caption = "  /  /  "
    Else
        DataUltChequeDevolvido.Caption = Format(objFilialClienteEstatistica.dtDataUltChequeDevolvido, "dd/mm/yyyy")
    End If

    '################################################
    'Inserido por Wagner 26/09/05
    If objFilialCliente.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objFilialCliente.iCodMensagem)
        Call Mensagem_Validate(bSGECancelDummy)
    End If
    '####################################

    iAlterado = 0

    Exibe_Dados_FilialCliente = SUCESSO

Exit Function

Erro_Exibe_Dados_FilialCliente:

    Exibe_Dados_FilialCliente = Err

    Select Case Err

        Case 6926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCodigoNome.lCodigo)

        Case 12557, 12558, 12429

        Case 12560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objFilialCliente.sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155840)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim colEndereco As New Collection
Dim colCategoriaItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Codigo do Cliente
    If Len(Trim(CodCliente.Text)) = 0 Then gError 12612

    'Verifica se foi preenchido Codigo da Filial
    If Len(Trim(CodFilial.Text)) = 0 Then gError 12613

    'Verifica se foi preenchido o Nome da Filial
    If Len(Trim(Nome.Text)) = 0 Then gError 12614

'    'Verifica se foi preenchido o Estado
'    If Len(Trim(Endereco(0).Text)) <> 0 Then
'        If Len(Trim(Estado(0).Text)) = 0 Then gError 43295
'    End If
'
'    If Len(Trim(Endereco(1).Text)) <> 0 Then
'        If Len(Trim(Estado(1).Text)) = 0 Then gError 43296
'    End If
'
'    If Len(Trim(Endereco(2).Text)) <> 0 Then
'        If Len(Trim(Estado(2).Text)) = 0 Then gError 43297
'    End If

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 33950
    Next

    'Lê os Enderecos e coloca em colEndereco
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEndereco)
    If lErro <> SUCESSO Then gError 12615

    'Lê os dados da tela relativos à Filial Cliente
    lErro = Le_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 12618
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objFilialCliente.sCgc)) > 0 Then
        'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
        lErro = CF("FilialCliente_Testa_CGC", objFilialCliente.lCodCliente, 0, objFilialCliente.sCgc)
        If lErro <> SUCESSO Then gError 58616
    End If
    
    'verifica preenchimento de campos de acordo com configuracao
    lErro = CF("FilialCliente_Verifica_Preenchimento", objFilialCliente, colEndereco.Item(1))
    If lErro <> SUCESSO Then gError 124068

    lErro = CF("InscricaoEstadual_Critica", objFilialCliente.sInscricaoEstadual, colEndereco.Item(1).sSiglaEstado)
    If lErro <> SUCESSO Then gError 189575

    lErro = Trata_Alteracao(objFilialCliente, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> SUCESSO Then gError 130398
        
    'Grava FilialCliente no BD
    lErro = CF("FiliaisClientes_Grava", objFilialCliente, colEndereco)
    If lErro <> SUCESSO Then gError 12619

'    'Altera a definição da Filial Cliente na árvore (TreeView)
'    lErro = Arvore_Alterar(Filiais, objFilialCliente)
'    If lErro <> SUCESSO And lErro <> 12716 Then gError 12714
'
'    'Se a filial não estava na árvore
'    If lErro = 12716 Then
'
'        'Inserir a FilialCliente na árvore
'        Call Arvore_Inserir(Filiais, objFilialCliente)
'
'    End If

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 12612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 12613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 12614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEFILIAL_NAO_PREENCHIDO", gErr)

        Case 12615, 12618, 12619, 12714, 58616, 124068, 130398, 189575

        Case 33950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", gErr)

        Case 43295, 43296, 43297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155841)

    End Select

    Exit Function

End Function

Private Function Arvore_Alterar(Filiais As TreeView, objFilialCliente As ClassFilialCliente) As Long
'Altera o nome da filial na árvore de Filiais

Dim lErro As Long
Dim objNode As Node
Dim sKeyFilial As String
Dim iAchou As Integer

On Error GoTo Erro_Arvore_Alterar

    sKeyFilial = KEY_CARACTER & CStr(objFilialCliente.lCodCliente) & SEPARADOR & CStr(objFilialCliente.iCodFilial)

    iAchou = 0

    'Pesquisa o nó da FilialCliente na árvore
    For Each objNode In Filiais.Nodes

        'Se encontrou
        If objNode.Key = sKeyFilial Then

            'Carrega o nó com o nome da Filial
            objNode.Text = objFilialCliente.sNome
            iAchou = 1

            Exit For

        End If

    Next

    'Se não achou a filial na árvore
    If iAchou = 0 Then Error 12716

    Arvore_Alterar = SUCESSO

    Exit Function

Erro_Arvore_Alterar:

    Arvore_Alterar = Err

    Select Case Err

        Case 12716  'Não encontrou na árvore

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155842)

    End Select

    Exit Function

End Function

Private Sub Arvore_Inserir(Filiais As TreeView, objFilialCliente As ClassFilialCliente)
'Insere a Filial Cliente na árvore

Dim objNode As Object
Dim lErro As Long, colNodes As Nodes
Dim sKeyFilial As String, objCliente As New ClassCliente
Dim sKeyPai As String, iAchou As Integer

On Error GoTo Erro_Arvore_Inserir

    'Chave do pai (Cliente)
    sKeyPai = KEY_CARACTER & CStr(objFilialCliente.lCodCliente)

    Set colNodes = Filiais.Nodes
    
    iAchou = 0

    'Pesquisa o pai (Cliente) na árvore
    For Each objNode In colNodes

        If objNode.Key = sKeyPai Then

            iAchou = 1
            Exit For

        End If

    Next

    'se não achou o Cliente na árvore
    If iAchou = 0 Then
    
        'obtem o nome reduzido do cliente
        objCliente.lCodigo = objFilialCliente.lCodCliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 56716
        
        'Não encontrou --> Erro
        If lErro <> SUCESSO Then Error 56717
        
        'inclui o cliente na arvore
        Set objNode = colNodes.Add(, tvwLast, sKeyPai)
        objNode.Text = objCliente.sNomeReduzido
    
    End If
    
    sKeyFilial = KEY_CARACTER & CStr(objFilialCliente.lCodCliente) & SEPARADOR & CStr(objFilialCliente.iCodFilial)
    
    'Adiciona o nó da Filial como filho do nó Cliente
    Call Filiais.Nodes.Add(sKeyPai, tvwChild, sKeyFilial, objFilialCliente.sNome)

    Exit Sub
    
Erro_Arvore_Inserir:
     
    Select Case Err
          
        Case 56716
        
        Case 56717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155843)
     
    End Select
     
    Exit Sub
    
End Sub

Private Sub Arvore_Excluir(Filiais As TreeView, objFilialCliente As ClassFilialCliente)

Dim objNode As Node
Dim sFilialCliente As String

    sFilialCliente = KEY_CARACTER & CStr(objFilialCliente.lCodCliente) & SEPARADOR & CStr(objFilialCliente.iCodFilial)

    'Pesquisa nó da FilialCliente na árvore
    For Each objNode In Filiais.Nodes
        
        
        'Se encontrou,
        If objNode.Key = sFilialCliente Then
            
            'Exclui nó da FilialCliente
            Filiais.Nodes.Remove sFilialCliente
            Exit For
        
        End If
    
    Next

    Exit Sub

End Sub

Private Function Le_Dados_FilialCliente(objFilialCliente As ClassFilialCliente) As Long
'Lê os dados que estão na tela de FiliaisClientes e coloca-os em objFilialCliente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sConta As String
Dim iContaPreenchida As Integer
Dim objVendedor As New ClassVendedor
Dim objFilialCliCategoria As New ClassFilialCliCategoria
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objCobrador As New ClassCobrador
Dim objTransportadora As New ClassTransportadora
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Le_Dados_FilialCliente

    'IDENTIFICACAO :
    
    objFilialCliente.iAtivo = Ativo.Value
    
    If Len(Trim(CodCliente.Text)) > 0 Then objFilialCliente.lCodCliente = CLng(CodCliente.Text)

    If Len(Trim(CodFilial.Text)) > 0 Then objFilialCliente.iCodFilial = CodFilial.Text
    
    objFilialCliente.sRG = Trim(RG.Text)
    objFilialCliente.sNome = Trim(Nome.Text)
    objFilialCliente.sCgc = Trim(CGC.Text)
    objFilialCliente.sInscricaoEstadual = Trim(InscricaoEstadual.Text)
    objFilialCliente.sInscricaoMunicipal = Trim(InscricaoMunicipal.Text)
    objFilialCliente.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)
    objFilialCliente.sObservacao = Trim(Observacao.Text)
    objFilialCliente.iRegimeTributario = RegimeTributario.ItemData(RegimeTributario.ListIndex)
    
    If IEIsento.Value = vbChecked Then
        objFilialCliente.iIEIsento = MARCADO
    Else
        objFilialCliente.iIEIsento = DESMARCADO
    End If
    If IENaoContrib.Value = vbChecked Then
        objFilialCliente.iIENaoContrib = MARCADO
    Else
        objFilialCliente.iIENaoContrib = DESMARCADO
    End If

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialCliCategoria = New ClassFilialCliCategoria

            objFilialCliCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialCliCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            'Verifica se já existe a categoria na coleção
            If objFilialCliente.colCategoriaItem.Count > 0 Then

                For iIndice1 = 1 To objFilialCliente.colCategoriaItem.Count

                    If objFilialCliCategoria.sCategoria = objFilialCliente.colCategoriaItem.Item(iIndice1).sCategoria Then Error 33092

                Next

            End If

            objFilialCliente.colCategoriaItem.Add objFilialCliCategoria

        End If

    Next

   'VENDAS :

    If Len(Trim(Vendedor.Text)) > 0 Then

        objVendedor.sNomeReduzido = Vendedor.Text

        'Tenta ler Vendedor a partir do NomeReduzido
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then Error 25037

        'Não encontrou o Vendedor
        If lErro <> SUCESSO Then Error 25038

        objFilialCliente.iVendedor = objVendedor.iCodigo

    End If

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12620

    If iContaPreenchida = CONTA_VAZIA Then
        objFilialCliente.sContaContabil = ""
    Else
        objFilialCliente.sContaContabil = sConta
    End If

    If Len(Trim(ComissaoVendas.Text)) > 0 Then objFilialCliente.dComissaoVendas = CDbl(ComissaoVendas.Text) / 100

    If Len(Trim(Regiao.Text)) > 0 Then objFilialCliente.iRegiao = Codigo_Extrai(Regiao.Text)

    'Verifica se região venda existe
    If objFilialCliente.iRegiao <> 0 Then
        objRegiaoVenda.iCodigo = objFilialCliente.iRegiao
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 33095

        'Não encontrou a região venda ==> erro
        If lErro = 16137 Then Error 33096

    End If

    If Len(Trim(Cobrador.Text)) > 0 Then objFilialCliente.iCodCobrador = Codigo_Extrai(Cobrador.Text)
    
    'Verifica se cobrador existe
    If objFilialCliente.iCodCobrador <> 0 Then
        objCobrador.iCodigo = objFilialCliente.iCodCobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 33097

        'Não encontrou o cobrador ==> erro
        If lErro = 19294 Then Error 33098

    End If

    If Len(Trim(FreqVisitas.Text)) > 0 Then objFilialCliente.iFreqVisitas = CInt(FreqVisitas.Text)

    If Trim(DataUltVisita.Text) = "/  /" Then
        objFilialCliente.dtDataUltVisita = DATA_NULA
    Else
        objFilialCliente.dtDataUltVisita = CDate(DataUltVisita.Text)
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then objFilialCliente.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    
    If TipoFrete.ListIndex <> -1 Then
        objFilialCliente.iTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    Else
        objFilialCliente.iTipoFrete = -1
    End If
    
    
    'Verifica se transportadora existe
    If objFilialCliente.iCodTransportadora <> 0 Then
        
        objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 33099

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then Error 33100

    End If

    objFilialCliente.sGuia = Trim(Guia.Text)
    
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objFilialCliente.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)

    'Verifica se transportadora existe
    If objFilialCliente.iCodTranspRedesp <> 0 Then
        objTransportadora.iCodigo = objFilialCliente.iCodTranspRedesp
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 33099

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then Error 33100

    End If
    
    objFilialCliente.iDetPagFrete = RedespachoCli.Value
    
    '#######################################
    'Inserido por Wagner 26/09/05
    If Len(Trim(Mensagem.Text)) > 0 Then objFilialCliente.iCodMensagem = Codigo_Extrai(Mensagem.Text)

    'Verifica se mensagem existe
    If objFilialCliente.iCodMensagem <> 0 Then
        objMensagem.iCodigo = objFilialCliente.iCodMensagem
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then Error 33017

        'Não encontrou a mensagem ==> erro
        If lErro = 19234 Then Error 33018

    End If
    '#######################################
    Le_Dados_FilialCliente = SUCESSO

    Exit Function

Erro_Le_Dados_FilialCliente:

    Le_Dados_FilialCliente = Err

    Select Case Err

        Case 12620, 25037, 33095, 33097, 33099, 33017
            'Erro tratado na rotina chamada

        Case 25038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)

        Case 33092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_REPETIDA_NO_GRID", Err)

        Case 33096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, objRegiaoVenda.iCodigo)

        Case 33098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)

        Case 33100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 33009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", Err)

        Case 33018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", Err, objMensagem.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155845)

    End Select

    Exit Function

End Function

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE VENDEDOR :

Public Sub VendedorLabel_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o vendedor estiver preenchido passa seu codigo para o objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os clientes
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

'BROWSE PLANO_CONTA :

Public Sub ContaContabilLabel_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12885
    
    'Se a conta estiver preenchida passa para objPlanoConta
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta
    
    'Chama a tela que lista as contas
    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 12885

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155847)

    End Select

    Exit Sub

End Sub

'RETORNO VENDEDORES :

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim bCancel As Boolean
Dim lErro As Long

    Set objVendedor = obj1

    Vendedor.Text = CStr(objVendedor.iCodigo)

    Call Vendedor_Validate(bCancel)

    Me.Show

    Vendedor.SetFocus 'Inserido por Wagner

    Exit Sub

End Sub

'RETORNO PLANO_CONTA :

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
        If lErro <> SUCESSO Then Error 12506

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 12506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155848)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA  "
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FiliaisClientes"

    'Lê os dados da Tela FilialCliente
    lErro = Le_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then Error 12603

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodCliente", objFilialCliente.lCodCliente, 0, "CodCliente"
    colCampoValor.Add "CodFilial", objFilialCliente.iCodFilial, 0, "CodFilial"
    colCampoValor.Add "Nome", objFilialCliente.sNome, STRING_FILIAL_CLIENTE_NOME, "Nome"
    colCampoValor.Add "RG", objFilialCliente.sRG, STRING_RG, "RG"
    colCampoValor.Add "CGC", objFilialCliente.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objFilialCliente.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "InscricaoMunicipal", objFilialCliente.sInscricaoMunicipal, STRING_INSCR_MUN, "InscricaoMunicipal"
    colCampoValor.Add "InscricaoSuframa", objFilialCliente.sInscricaoSuframa, STRING_INSCR_SUF, "InscricaoSuframa"
    colCampoValor.Add "Endereco", objFilialCliente.lEndereco, 0, "Endereco"
    colCampoValor.Add "EnderecoEntrega", objFilialCliente.lEnderecoEntrega, 0, "EnderecoEntrega"
    colCampoValor.Add "EnderecoCobranca", objFilialCliente.lEnderecoCobranca, 0, "EnderecoCobranca"
    colCampoValor.Add "CodTransportadora", objFilialCliente.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objFilialCliente.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objFilialCliente.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "Guia", objFilialCliente.sGuia, STRING_CLIENTE_GUIA, "Guia"
    colCampoValor.Add "Observacao", objFilialCliente.sObservacao, STRING_CLIENTE_OBSERVACAO, "Observacao"
    colCampoValor.Add "ContaContabil", objFilialCliente.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "Vendedor", objFilialCliente.iVendedor, 0, "Vendedor"
    colCampoValor.Add "ComissaoVendas", objFilialCliente.dComissaoVendas, 0, "ComissaoVendas"
    colCampoValor.Add "Regiao", objFilialCliente.iRegiao, 0, "Regiao"
    colCampoValor.Add "FreqVisitas", objFilialCliente.iFreqVisitas, 0, "FreqVisitas"
    colCampoValor.Add "DataUltVisita", objFilialCliente.dtDataUltVisita, 0, "DataUltVisita"
    colCampoValor.Add "CodCobrador", objFilialCliente.iCodCobrador, 0, "CodCobrador"
    colCampoValor.Add "TipoFrete", objFilialCliente.iTipoFrete, 0, "TipoFrete"
    colCampoValor.Add "Ativo", objFilialCliente.iAtivo, 0, "Ativo"
    colCampoValor.Add "CodMensagem", objFilialCliente.iCodMensagem, 0, "CodMensagem" 'Inserido por Wagner '18/11/2005
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12603

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155849)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Tela_Preenche

    objFilialCliente.lCodCliente = colCampoValor.Item("CodCliente").vValor

    'Se código de Cliente estiver preenchido,
    If objFilialCliente.lCodCliente <> 0 Then

        'Passa os dados da coleção para o objeto
        objFilialCliente.iCodFilial = colCampoValor.Item("CodFilial").vValor
        objFilialCliente.sNome = colCampoValor.Item("Nome").vValor
        objFilialCliente.sCgc = colCampoValor.Item("CGC").vValor
        objFilialCliente.sRG = colCampoValor.Item("RG").vValor
        objFilialCliente.sInscricaoEstadual = colCampoValor.Item("InscricaoEstadual").vValor
        objFilialCliente.sInscricaoMunicipal = colCampoValor.Item("InscricaoMunicipal").vValor
        objFilialCliente.sInscricaoSuframa = colCampoValor.Item("InscricaoSuframa").vValor
        objFilialCliente.lEndereco = colCampoValor.Item("Endereco").vValor
        objFilialCliente.lEnderecoEntrega = colCampoValor.Item("EnderecoEntrega").vValor
        objFilialCliente.lEnderecoCobranca = colCampoValor.Item("EnderecoCobranca").vValor
        objFilialCliente.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
        objFilialCliente.iCodTranspRedesp = colCampoValor.Item("CodTranspRedesp").vValor
        objFilialCliente.iDetPagFrete = colCampoValor.Item("DetPagFrete").vValor
        objFilialCliente.sGuia = colCampoValor.Item("Guia").vValor
        objFilialCliente.sObservacao = colCampoValor.Item("Observacao").vValor
        objFilialCliente.sContaContabil = colCampoValor.Item("ContaContabil").vValor
        objFilialCliente.iVendedor = colCampoValor.Item("Vendedor").vValor
        objFilialCliente.dComissaoVendas = colCampoValor.Item("ComissaoVendas").vValor
        objFilialCliente.iRegiao = colCampoValor.Item("Regiao").vValor
        objFilialCliente.iFreqVisitas = colCampoValor.Item("Freqvisitas").vValor
        objFilialCliente.dtDataUltVisita = colCampoValor.Item("DataUltVisita").vValor
        objFilialCliente.iCodCobrador = colCampoValor.Item("CodCobrador").vValor
        objFilialCliente.iTipoFrete = colCampoValor.Item("TipoFrete").vValor
        objFilialCliente.iAtivo = colCampoValor.Item("Ativo").vValor
        objFilialCliente.iCodMensagem = colCampoValor.Item("CodMensagem").vValor 'Inserido por Wagner
        
        objFilialClienteEstatistica.lCodCliente = objFilialCliente.lCodCliente
        objFilialClienteEstatistica.iFilialEmpresa = giFilialEmpresa
        objFilialClienteEstatistica.iCodFilial = objFilialCliente.iCodFilial
        
        'Tenta ler FilialCliente
        lErro = CF("FilialCliente_Le_Estendida", objFilialCliente, objFilialClienteEstatistica)
        If lErro <> SUCESSO And lErro <> 52591 Then gError 52700
        
        'Exibe os dados da FilialCliente
        lErro = Exibe_Dados_FilialCliente(objFilialCliente, objFilialClienteEstatistica)
        If lErro <> SUCESSO Then gError 12601
        
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 12601, 52700

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155850)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

 Dim lErro As Long

    Set objEventoVendedor = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoCobrador = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoCliente = Nothing
    Set objEventoFilialCliente = Nothing
    
     Set objEventoMensagem = Nothing 'Inserido por Wagner 26/09/05
     
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set objGridCategoria = Nothing

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCategoria.Col

            Case GRID_CATEGORIA_COL

                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then Error 33082

            Case GRID_VALOR_COL

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then Error 33083

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 33084

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 33082, 33083

        Case 33084
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155851)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaCliente

    iIndice = GridCategoria.Row
    If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) <> ComboCategoriaCliente.Text Then

        If Len(Trim(ComboCategoriaCliente.Text)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaCliente.sCategoria = ComboCategoriaCliente.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaCliente_Le", objCategoriaCliente)
            If lErro <> SUCESSO And lErro <> 28847 Then Error 33085

            'Categoria não está cadastrada
            If lErro = 28847 Then Error 33086

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaCliente.Text Then Error 33087

                    End If

                Next

            End If

            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                'Limpa o valor (ítem) da categoria
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        Else
            
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row
               
                'Limpa o valor (ítem) da categoria
                GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL) = ""

            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33088

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = Err

    Select Case Err

        Case 33085
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33086 'Categoria não está cadastrada
            
            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaCliente"
                Call Chama_Tela("CategoriaCliente", objCategoriaCliente)

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 33087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33088
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155852)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colItens As New Collection
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaClienteItem

    If Len(Trim(ComboCategoriaClienteItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then Error 55908

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaClienteItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaClienteItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaClienteItem.sItem = ComboCategoriaClienteItem.Text

            'Lê Item De Categoria De Cliente no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaClienteItem)
            If lErro <> SUCESSO And lErro <> 28991 Then Error 33089

            'Item da Categoria não está cadastrado
            If lErro = 28991 Then Error 33090

        End If

    End If

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33091

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 33089
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33090 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTEITEM")

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaCliente.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

                'Chama a Tela "CategoriaCliente"
                Call Chama_Tela("CategoriaCliente", objCategoriaCliente, objCategoriaClienteItem.sItem)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 33091
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55908
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155853)

    End Select

    Exit Function

End Function

Private Sub objEventoCobrador_evSelecao(obj1 As Object)

Dim objCobrador As ClassCobrador
Dim bCancel As Boolean

    Set objCobrador = obj1

    'Preenche campo Agente Cobrador
    Cobrador.Text = objCobrador.iCodigo

    Call Cobrador_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ID
    Set Form_Load_Ocx = Me
    Caption = "Filiais de Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FiliaisClientes"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Cobrador Then
            Call LabelCobrador_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is CodCliente Then
            Call LabelCodCliente_Click
        ElseIf Me.ActiveControl Is CodFilial Then
            Call LabelCodFilial_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is Mensagem Then 'Inserido por Wagner 26/09/05
            Call MensagemNFLabel_Click
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

Public Sub TranspRedespacho_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TranspRedespacho_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TranspRedespacho_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_TranspRedespacho_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(TranspRedespacho.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If TranspRedespacho.Text = TranspRedespacho.List(TranspRedespacho.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TranspRedespacho, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19243

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 19244
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then Error 19245
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19246

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19243, 19244

        Case 19245  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 19246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155854)

    End Select

    Exit Sub

End Sub

Public Sub TranspRedLabel_Click()
'BROWSE TRANSPORTADORA :

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(TranspRedespacho.Text)
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTranspRedespacho)

End Sub

Private Sub objEventoTranspRedespacho_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora
Dim bCancel As Boolean

    Set objTransportadora = obj1

    'Preenche campo Transportadora
    TranspRedespacho.Text = objTransportadora.iCodigo

    'Chama a rotina de validate
    Call TranspRedespacho_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Guia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RedespachoCli_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Mensagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Mensagem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Mensagem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMensagem As New ClassMensagem
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Mensagem_Validate

    'Verifica se foi preenchida a ComboBox Mensagem
    If Len(Trim(Mensagem.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Mensagem
    If Mensagem.Text = Mensagem.List(Mensagem.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Mensagem, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19227

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objMensagem.iCodigo = iCodigo

        'Tenta ler Mensagem com esse código no BD
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then Error 19228

        If lErro <> SUCESSO Then Error 19229 'Não encontrou Mensagem no BD

        'Encontrou Mensagem no BD, coloca no Text da Combo
        Mensagem.Text = CStr(objMensagem.iCodigo) & SEPARADOR & objMensagem.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19230

    Exit Sub

Erro_Mensagem_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19227, 19228

    Case 19229  'Não encontrou Mensagem no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MENSAGEM")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Mensagens", objMensagem)

        End If

    Case 19230

        lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_ENCONTRADA", Err, Mensagem.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155857)

    End Select

    Exit Sub

End Sub

Public Sub MensagemNFLabel_Click()

'BROWSE MENSAGEM:

Dim objMensagem As New ClassMensagem
Dim colSelecao As New Collection
    
    'Se a Mensagem estiver preenchida então passa para o objMensagem o Código da Mensagem
    If Len(Trim(Mensagem.Text)) > 0 Then objMensagem.iCodigo = Codigo_Extrai(Mensagem.Text)
    
    'Chama a tela que lista todas as Mensagens
    Call Chama_Tela("MensagemLista", colSelecao, objMensagem, objEventoMensagem)

End Sub

Private Sub objEventoMensagem_evSelecao(obj1 As Object)

Dim objMensagem As ClassMensagem
Dim bCancel As Boolean

    Set objMensagem = obj1

    'Preenche campo Mensagem p/ NF
    Mensagem.Text = objMensagem.iCodigo

    'Chama a rotina de validate
    Call Mensagem_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207446

    TipoFrete.Clear

    For Each objTipoFrete In colTipoFrete

        'coloca na combo
        TipoFrete.AddItem objTipoFrete.sDescricao
        TipoFrete.ItemData(TipoFrete.NewIndex) = objTipoFrete.iCodigoNFE

    Next

    Carrega_TipoFrete = SUCESSO

    Exit Function

Erro_Carrega_TipoFrete:

    Carrega_TipoFrete = gErr

    Select Case gErr

        'Erro já tratado
        Case 207446

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207447)

    End Select

    Exit Function

End Function

Private Function Inicializa_Tab_Inscricao() As Long

Dim lErro  As Long

On Error GoTo Erro_Inicializa_Tab_Inscricao

    RegimeTributario.AddItem REGIME_TRIBUTARIO_NORMAL_TEXTO
    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_NORMAL
    RegimeTributario.AddItem REGIME_TRIBUTARIO_SIMPLES_TEXTO
    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_SIMPLES

    RegimeTributario.ListIndex = 0
    
    Inicializa_Tab_Inscricao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Inscricao:
    
    Inicializa_Tab_Inscricao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207400)
        
    End Select
        
    Exit Function
            
End Function

Public Sub IEIsento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IENaoContrib_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Trata_IE()
    If Len(Trim(InscricaoEstadual.Text)) > 0 Then
        IEIsento.Value = vbUnchecked
        IEIsento.Enabled = False
        If InscricaoEstadual.Text <> sIEAnt Then
            IENaoContrib.Value = vbUnchecked
        End If
    Else
        If InscricaoEstadual.Text <> sIEAnt Then
            IEIsento.Value = vbChecked
            IENaoContrib.Value = vbChecked
        End If
        IEIsento.Enabled = True
    End If
    sIEAnt = InscricaoEstadual.Text
End Sub

Public Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call Trata_IE
End Sub
