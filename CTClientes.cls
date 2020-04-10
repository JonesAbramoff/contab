VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjInfoUsu As Object 'Inserido por Wagner
Dim gcolRotas As New Collection

''''==================== ANOTACAO =====================
Public gobjAnotacao As ClassAnotacoes
''''===================================================
Public gobjTabEnd As Object

Dim iFrameAtual As Integer
Dim iFrame2Atual As Integer
Public iAlterado As Integer
Dim iVendedorAlterado As Integer
Dim sIEAnt As String

Dim objGridComissao As AdmGrid

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Dim objGridRotas As AdmGrid
Dim iGrid_RotaFilial_Col As Integer
Dim iGrid_RotaChave_Col As Integer
Dim iGrid_RotaCodigo_Col As Integer

Private WithEvents objEventoTipoCliente As AdmEvento
Attribute objEventoTipoCliente.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCobrador As AdmEvento
Attribute objEventoCobrador.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoPadraoCobranca As AdmEvento
Attribute objEventoPadraoCobranca.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

'Private WithEvents objEventoClienteFilial As AdmEvento

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DadosFinanceiros = 2
Private Const TAB_Inscricoes = 3
Private Const TAB_Enderecos = 4
Private Const TAB_Vendas = 5
Private Const TAB_Estatisticas = 6

'--- inicio dos properties get dos controles da tela
Public Property Get IEIsento() As Object
     Set IEIsento = objUserControl.Controls("IEIsento")
End Property

Public Property Get IENaoContrib() As Object
     Set IENaoContrib = objUserControl.Controls("IENaoContrib")
End Property

Public Property Get TotalCR() As Object
     Set TotalCR = objUserControl.Controls("TotalCR")
End Property

Public Property Get PercCREmAberto() As Object
     Set PercCREmAberto = objUserControl.Controls("PercCREmAberto")
End Property

Public Property Get TotalCREmCartorio() As Object
     Set TotalCREmCartorio = objUserControl.Controls("TotalCREmCartorio")
End Property

Public Property Get PercCREmCartorio() As Object
     Set PercCREmCartorio = objUserControl.Controls("PercCREmCartorio")
End Property

Public Property Get TotalCRComProtesto() As Object
     Set TotalCRComProtesto = objUserControl.Controls("TotalCRComProtesto")
End Property

Public Property Get PercCRComProtesto() As Object
     Set PercCRComProtesto = objUserControl.Controls("PercCRComProtesto")
End Property

'####################################################
'Inserido por Wagner
Public Property Get Bloqueado() As Object
     Set Bloqueado = objUserControl.Controls("Bloqueado")
End Property
'####################################################

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

Public Property Get PadraoCobranca() As Object
     Set PadraoCobranca = objUserControl.Controls("PadraoCobranca")
End Property

Public Property Get Ativo() As Object
     Set Ativo = objUserControl.Controls("Ativo")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get Cobrador() As Object
     Set Cobrador = objUserControl.Controls("Cobrador")
End Property

Public Property Get Regiao() As Object
     Set Regiao = objUserControl.Controls("Regiao")
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

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get Label1() As Object
     Set Label1 = objUserControl.Controls("Label1")
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get PadraoCobrancaLabel() As Object
     Set PadraoCobrancaLabel = objUserControl.Controls("PadraoCobrancaLabel")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get VendedorLabel() As Object
     Set VendedorLabel = objUserControl.Controls("VendedorLabel")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get AgenteCobradorLabel() As Object
     Set AgenteCobradorLabel = objUserControl.Controls("AgenteCobradorLabel")
End Property

Public Property Get NumeroCompras() As Object
     Set NumeroCompras = objUserControl.Controls("NumeroCompras")
End Property

Public Property Get MediaCompra() As Object
     Set MediaCompra = objUserControl.Controls("MediaCompra")
End Property

Public Property Get ValorAcumuladoCompras() As Object
     Set ValorAcumuladoCompras = objUserControl.Controls("ValorAcumuladoCompras")
End Property

Public Property Get DataPrimeiraCompra() As Object
     Set DataPrimeiraCompra = objUserControl.Controls("DataPrimeiraCompra")
End Property

Public Property Get DataUltimaCompra() As Object
     Set DataUltimaCompra = objUserControl.Controls("DataUltimaCompra")
End Property

Public Property Get NumChequesDevolvidos() As Object
     Set NumChequesDevolvidos = objUserControl.Controls("NumChequesDevolvidos")
End Property

Public Property Get DataUltChequeDevolvido() As Object
     Set DataUltChequeDevolvido = objUserControl.Controls("DataUltChequeDevolvido")
End Property

Public Property Get SaldoAtrasados() As Object
     Set SaldoAtrasados = objUserControl.Controls("SaldoAtrasados")
End Property

Public Property Get ValorPagtosAtraso() As Object
     Set ValorPagtosAtraso = objUserControl.Controls("ValorPagtosAtraso")
End Property

Public Property Get MediaAtraso() As Object
     Set MediaAtraso = objUserControl.Controls("MediaAtraso")
End Property

Public Property Get MaiorAtraso() As Object
     Set MaiorAtraso = objUserControl.Controls("MaiorAtraso")
End Property

Public Property Get SaldoTitulos() As Object
     Set SaldoTitulos = objUserControl.Controls("SaldoTitulos")
End Property

Public Property Get SaldoDuplicatas() As Object
     Set SaldoDuplicatas = objUserControl.Controls("SaldoDuplicatas")
End Property

Public Property Get SaldoPedidosLiberados() As Object
     Set SaldoPedidosLiberados = objUserControl.Controls("SaldoPedidosLiberados")
End Property

Public Property Get OpcaoEndereco() As Object
     Set OpcaoEndereco = objUserControl.Controls("OpcaoEndereco")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get CondicaoPagto() As Object
     Set CondicaoPagto = objUserControl.Controls("CondicaoPagto")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get ComboCategoriaCliente() As Object
     Set ComboCategoriaCliente = objUserControl.Controls("ComboCategoriaCliente")
End Property

Public Property Get ComboCategoriaClienteItem() As Object
     Set ComboCategoriaClienteItem = objUserControl.Controls("ComboCategoriaClienteItem")
End Property

Public Property Get GridCategoria() As Object
     Set GridCategoria = objUserControl.Controls("GridCategoria")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get RazaoSocial() As Object
     Set RazaoSocial = objUserControl.Controls("RazaoSocial")
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objUserControl.Controls("NomeReduzido")
End Property

Public Property Get LimiteCredito() As Object
     Set LimiteCredito = objUserControl.Controls("LimiteCredito")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get CondicaoPagtoLabel() As Object
     Set CondicaoPagtoLabel = objUserControl.Controls("CondicaoPagtoLabel")
End Property

Public Property Get MensagemNFLabel() As Object
     Set MensagemNFLabel = objUserControl.Controls("MensagemNFLabel")
End Property

Public Property Get TipoClienteLabel() As Object
     Set TipoClienteLabel = objUserControl.Controls("TipoClienteLabel")
End Property

Public Property Get Observacao2() As Object
     Set Observacao2 = objUserControl.Controls("Observacao2")
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

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get Filiais() As Object
     Set Filiais = objUserControl.Controls("Filiais")
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

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get ComboCobrador() As Object
     Set ComboCobrador = objUserControl.Controls("ComboCobrador")
End Property

Public Property Get UsuRespCallCenter() As Object
     Set UsuRespCallCenter = objUserControl.Controls("UsuRespCallCenter")
End Property

Public Property Get SaldoLimitedeCredito() As Object
     Set SaldoLimitedeCredito = objUserControl.Controls("SaldoLimitedeCredito")
End Property

Public Property Get SaldodeCredito() As Object
     Set SaldodeCredito = objUserControl.Controls("SaldodeCredito")
End Property

Public Property Get RecebForaFaixa() As Object
     Set RecebForaFaixa = objUserControl.Controls("RecebForaFaixa")
End Property

Public Property Get PercentMaisReceb() As Object
     Set PercentMaisReceb = objUserControl.Controls("PercentMaisReceb")
End Property

Public Property Get PercentMenosReceb() As Object
     Set PercentMenosReceb = objUserControl.Controls("PercentMenosReceb")
End Property

Public Property Get NaoTemFaixaReceb() As Object
     Set NaoTemFaixaReceb = objUserControl.Controls("NaoTemFaixaReceb")
End Property

Public Property Get IgnoraRecebPadrao() As Object
     Set IgnoraRecebPadrao = objUserControl.Controls("IgnoraRecebPadrao")
End Property

Public Property Get TabEnd() As Object
     Set TabEnd = objUserControl.Controls("TabEnd")
End Property

Public Property Get GridRotas() As Object
     Set GridRotas = objUserControl.Controls("GridRotas")
End Property

Public Property Get RotaFilial() As Object
     Set RotaFilial = objUserControl.Controls("RotaFilial")
End Property

Public Property Get RotaChave() As Object
     Set RotaChave = objUserControl.Controls("RotaChave")
End Property

Public Property Get RotaCodigo() As Object
     Set RotaCodigo = objUserControl.Controls("RotaCodigo")
End Property

Public Property Get FrameRotas() As Object
     Set FrameRotas = objUserControl.Controls("FrameRotas")
End Property

Public Property Get RotaFilialLista() As Object
     Set RotaFilialLista = objUserControl.Controls("RotaFilialLista")
End Property

Public Property Get RotaTodasFiliais() As Object
     Set RotaTodasFiliais = objUserControl.Controls("RotaTodasFiliais")
End Property

Public Property Get RegimeTributario() As Object
     Set RegimeTributario = objUserControl.Controls("RegimeTributario")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do próximo cliente
    lErro = CF("Cliente_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 57529

    'Exibe código na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57529
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155551)
    
    End Select

    Exit Sub

End Sub

Public Sub AgenteCobradorLabel_Click()

Dim objCobrador As New ClassCobrador
Dim colSelecao As New Collection
    
    'Se o Cobrador estiver preenchido --> passa o Código para o objCobrador
    If Len(Trim(Cobrador.Text)) > 0 Then objCobrador.iCodigo = Codigo_Extrai(Cobrador.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("CobradorLista", colSelecao, objCobrador, objEventoCobrador)

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 12430

    objCliente.lCodigo = CLng(Codigo.Text)

    'Lê os dados do cliente a ser excluido
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 12431

    'Verifica se cliente não está cadastrado
    If lErro = 12293 Then Error 12432

    'Envia aviso perguntando se realmente deseja excluir cliente e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CLIENTE", objCliente.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Cliente
        lErro = CF("Cliente_Exclui", objCliente)
        If lErro <> SUCESSO Then Error 12467

'        'Exclui da ListBox
'        Call ClientesList_Exclui(objCliente)

        'Limpa a Tela
        lErro = Limpa_Tela_Clientes()
        If lErro <> SUCESSO Then Error 58587

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 12431, 12467, 58587
        
        Case 12430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", Err)

        Case 12432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155552)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub RG_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RG_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19239

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 19240
        
        'Não encontrou Cobrador no BD
        If lErro <> SUCESSO Then Error 19241
        
        If giFilialEmpresa <> EMPRESA_TODA Then
            
            If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
                If objCobrador.iFilialEmpresa <> giFilialEmpresa Then Error 49556
            End If
            
        End If

        'Encontrou Cobrador no BD, coloca no Text da Combo
        Cobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19242

    Exit Sub

Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19239, 19240

        Case 19241  'Não encontrou Cobrador no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

            If vbMsgRes = vbYes Then
    
                 Call Chama_Tela("Cobradores", objCobrador)
    
            End If
    
        Case 19242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, Cobrador.Text)
   
        Case 49556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_PERTENCE_FILIAL", Err, objCobrador.iCodigo, giFilialEmpresa)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155553)
    
    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é um Long
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 19299

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 19299

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155554)

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
    Call Trata_ComboCategoriaClienteItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaClienteItem_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155555)

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

Public Sub CondicaoPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagto.Text = CondicaoPagto.List(CondicaoPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19223

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 19224
        
        'Não encontrou CondicaoPagto no BD
        If lErro <> SUCESSO Then Error 19225

        'Encontrou CondicaoPagto no BD e não é de Recebimento
        If objCondicaoPagto.iEmRecebimento = 0 Then Error 6927

        'Coloca no Text da Combo
        CondicaoPagto.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19226

    Exit Sub

Erro_CondicaoPagto_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 6927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)

        Case 19223, 19224

        Case 19225  'Não encontrou CondicaoPagto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If

        Case 19226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155556)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection
    
    'Se a codicao estiver preenchida então passa para o objCondicaoPagamento
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagto)
    
    'Chama tela que lista as Condições de Pagamento
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

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
    
    'Formata a Conta
    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 12261
    
    'Se a conta estiver preenchida --> passa para o objPlanoConta
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta
    
    'Chama a lista de Contas
    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 12261

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155557)

    End Select

    Exit Sub

End Sub

Public Sub DataUltVisita_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataUltVisita, iAlterado)

End Sub

Public Sub Filiais_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filiais_Click

    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 12486

    'Preenche objCliente
    objCliente.lCodigo = CLng(Codigo.Text)
    
    'Lê o Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 43495

    'Se não achou o Cliente
    If lErro <> SUCESSO Then

            'Envia aviso perguntando se deseja cadastrar novo Cliente
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE")

            If vbMsgRes = vbYes Then
                
                'Grava o novo cliente
                lErro = Gravar_Registro()
                If lErro <> SUCESSO Then Error 43496
                
                'Chama a Tela de Filiais de Cliente
                objFilialCliente.lCodCliente = CLng(Codigo.Text)
                objFilialCliente.iCodFilial = 1 'p/começar exibindo a matriz
                
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            
            End If
    Else
    
        'Chama a Tela de Filiais de Cliente
        objFilialCliente.lCodCliente = CLng(Codigo.Text)
        objFilialCliente.iCodFilial = 1 'p/começar exibindo a matriz
        
        Call Chama_Tela("FiliaisClientes", objFilialCliente)
    
    End If

    Exit Sub

Erro_Filiais_Click:

    Select Case Err

        Case 12486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", Err)

        Case 43495, 43496

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155558)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Cliente
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 12360

    'Limpa a Tela
    lErro = Limpa_Tela_Clientes()
    If lErro <> SUCESSO Then Error 58588
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 12360, 58588

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155559)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Cliente necessários foram preenchidos
'Grava Cliente no BD
'Atualiza ListBox de Clientes

Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente
Dim colEndereco As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 12361

    'Verifica se foi preenchida a Razao Social
    If Len(Trim(RazaoSocial.Text)) = 0 Then gError 12362

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 12363

'    If giTipoVersao = VERSAO_FULL Then
        'Verifica se foi preenchido o Tipo
        If Len(Trim(Tipo.Text)) = 0 Then gError 6902
'    End If

    'Verifica se foi preenchido o Estado dos Endereços
'    If Len(Trim(Endereco(0).Text)) <> 0 Then
'        If Len(Trim(Estado(0).Text)) = 0 Then gError 43290
'    End If
'
'    If Len(Trim(Endereco(1).Text)) <> 0 Then
'        If Len(Trim(Estado(1).Text)) = 0 Then gError 43291
'    End If
'
'    If Len(Trim(Endereco(2).Text)) <> 0 Then
'        If Len(Trim(Estado(2).Text)) = 0 Then gError 43292
'    End If

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
         If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 12356
    Next
    
    lErro = CF2(Me, "CLI_Grava_Registro")
    If lErro <> SUCESSO Then gError 192965

'    'Lê os dados dos 3 três Enderecos e coloca em colEndereco
'    lErro = Le_Dados_Enderecos(colEndereco)
'    If lErro <> SUCESSO Then gError 12507

    lErro = gobjTabEnd.Move_Endereco_Memoria(colEndereco)
    If lErro <> SUCESSO Then gError 12507

    'Lê os dados da Tela relacionados ao Cliente
    lErro = Le_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then gError 43293
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objCliente.sCgc)) > 0 Then
        'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
        lErro = CF("FilialCliente_Testa_CGC", objCliente.lCodigo, 0, objCliente.sCgc)
        If lErro <> SUCESSO Then gError 58615
    End If
    
    'verifica preenchimento de campos de acordo com configuracao
    lErro = CF("Cliente_Verifica_Preenchimento", objCliente, colEndereco.Item(1))
    If lErro <> SUCESSO Then gError 124067
    
    lErro = CF("InscricaoEstadual_Critica", objCliente.sInscricaoEstadual, colEndereco.Item(1).sSiglaEstado)
    If lErro <> SUCESSO Then gError 189576
    
    lErro = Trata_Alteracao(objCliente, objCliente.lCodigo)
    If lErro <> SUCESSO Then gError 130397

    lErro = CF2(Me, "CLI_Grava_Registro1", objCliente, colEndereco)
    If lErro <> SUCESSO Then gError 196228

    'Grava o Cliente no BD
    lErro = CF("Cliente_Grava", objCliente, colEndereco)
    If lErro <> SUCESSO Then gError 43294

'    'Atualiza ListBox de Clientes
'    Call ClientesList_Exclui(objCliente)
'    Call ClientesList_Inclui(objCliente)

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_CLIENTE
            gobjAnotacao.sID = CStr(objCliente.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 128148
            
        End If
        
    End If
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 6902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 12361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 12362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 12363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 12507, 43293, 43294, 58615, 124067, 130397, 189576, 192965 'Tratados nas Rotinas chamadas

        Case 12356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_VALOR_CORRESPONDENTE", gErr, GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))

        Case 43290, 43291, 43292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case 196228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155560)

    End Select

    Exit Function

End Function

'Public Sub ClientesList_Exclui(objCliente As ClassCliente)
''Percorre a ListBox Clienteslist para remover o cliente caso ele exista
'
'Dim iIndice As Integer
'
'    For iIndice = 0 To ClientesList.ListCount - 1
'
'        If ClientesList.ItemData(iIndice) = objCliente.lCodigo Then
'
'            ClientesList.RemoveItem iIndice
'            Exit For
'
'        End If
'
'    Next
'
'End Sub
'
'Public Sub ClientesList_Inclui(objCliente As ClassCliente)
''Inclui Cliente na List
'
'    ClientesList.AddItem objCliente.sNomeReduzido
'    ClientesList.ItemData(ClientesList.NewIndex) = objCliente.lCodigo
'
'End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 12484

    'Limpa a Tela
    lErro = Limpa_Tela_Clientes()
    If lErro <> SUCESSO Then Error 58589
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 12484, 58589

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155561)

    End Select

End Sub

Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12316
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12317
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            Error 12318

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case Err

        Case 12316, 12317

        Case 12318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155562)

    End Select


    Exit Sub

End Sub

'Public Sub ClientesList_DblClick()
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'Dim objClienteEstatistica As New ClassFilialClienteEst
'
'On Error GoTo Erro_ClientesList_DblClick
'
'    'Guarda o valor do código do Cliente selecionado na ListBox ClientesList
'    objCliente.lCodigo = ClientesList.ItemData(ClientesList.ListIndex)
'
'    'Lê o Cliente no BD
'    lErro = CF("Cliente_Le_Estendida",objCliente, objClienteEstatistica)
'    If lErro <> SUCESSO And lErro <> 52545 Then Error 12349
'
'    'Se cliente não está cadastrado, erro
'    If lErro = 12293 Then Error 52545
'
'    'Exibe os dados do Cliente
'    lErro = Exibe_Dados_Cliente(objCliente, objClienteEstatistica)
'    If lErro <> SUCESSO Then Error 12496
'
'    'Fecha o comando das setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Exit Sub
'
'Erro_ClientesList_DblClick:
'
'    Select Case Err
'
'        Case 12349, 12496
'
'        Case 52545
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)
'            ClientesList.RemoveItem (ClientesList.ListIndex)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155563)
'
'    End Select
'
'    ClientesList.SetFocus
'
'    Exit Sub
'
'End Sub

Public Sub Cobrador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoVendas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoVendas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ComissaoVendas_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(ComissaoVendas.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(ComissaoVendas.Text)
    If lErro <> SUCESSO Then Error 12334

    'Formata
    ComissaoVendas.Text = Format(ComissaoVendas.Text, "Fixed")

    Exit Sub

Erro_ComissaoVendas_Validate:

    Cancel = True


    Select Case Err

        Case 12334
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155564)
    
    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagto_Change()

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
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 44103

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 44104

        'Passa a conta para o campo na tela
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 12332

        'conta não cadastrada
        If lErro = 5700 Then Error 12333

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case Err

        Case 12332, 44103

        Case 12333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)

        Case 44104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155565)

    End Select


    Exit Sub

End Sub

Public Sub DataUltVisita_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataUltVisita_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataUltVisita_Validate

    'verifica se foi preenchida a Data
    If Len(Trim(DataUltVisita.ClipText)) = 0 Then Exit Sub

    'critica a data
    lErro = Data_Critica(DataUltVisita.Text)
    If lErro <> SUCESSO Then Error 12341

    Exit Sub

Erro_DataUltVisita_Validate:

    Cancel = True


    Select Case Err

        Case 12341

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155566)

    End Select

    Exit Sub

End Sub

Public Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDesconto As String

On Error GoTo Erro_Desconto_Validate

    sDesconto = Desconto.Text

    'verifica se foi preenchido o Desconto
    If Len(Trim(Desconto.Text)) = 0 Then Exit Sub

    'critica porcentagem
    lErro = Porcentagem_Critica(Desconto.Text)
    If lErro <> SUCESSO Then Error 12313

    Desconto.Text = Format(sDesconto, "Fixed")

    Exit Sub

Erro_Desconto_Validate:

    Cancel = True


    Select Case Err

        Case 12313

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155567)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objUser As Object
Dim objTela As Object

On Error GoTo Erro_Form_Load
    
    Ativo.Value = MARCADO
    
''''======================== ANOTACAO =========================
'''    Set gobjAnotacao = New ClassAnotacoes
''''===========================================================
    
    Call CF2(Me, "CLI_Trata_Parametros")
    
    'Inicializa variávies AdmEvento
    Set objEventoTipoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoCondicaoPagto = New AdmEvento
    Set objEventoCobrador = New AdmEvento
    Set objEventoMensagem = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoTranspRedespacho = New AdmEvento
    Set objEventoPadraoCobranca = New AdmEvento
    Set objEventoCliente = New AdmEvento

    'Inicializa variáveis de frame
    iFrame2Atual = 0
    iFrameAtual = 1

    RazaoSocial.MaxLength = STRING_CLIENTE_RAZAO_SOCIAL

    'Configura Endereço Principal como visível
    OpcaoEndereco(0).Value = True
    OpcaoEndereco(1).Value = False
    OpcaoEndereco(2).Value = False
    
    Set objTela = Me
    
    Set objGridRotas = New AdmGrid
    If gobjFAT.iExibeRotasCliente = MARCADO Then
        FrameRotas.Visible = True
        lErro = Inicializa_Grid_Rotas(objGridRotas)
        If lErro <> SUCESSO Then gError 58084
    End If

    'Prepara as Combos e Grid do Tab de Identificacao
    lErro = Inicializa_Tab_Identificacao()
    If lErro <> SUCESSO Then gError 58084
    
    'Prepara as Combos do Tab de Dados Financeiros
    lErro = Inicializa_Tab_DadosFinanceiros()
    If lErro <> SUCESSO Then gError 58085
    
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then gError 58086
    
    'Prepara as Combos e a Mascara da Conta Contabil do Tab de Vendas
    lErro = Inicializa_Tab_Vendas()
    If lErro <> SUCESSO Then gError 58087
        
'    If giTipoVersao = VERSAO_LIGHT Then
'        TabelaPreco.ListIndex = 0
'    End If
        
    Set objUser = objUserControl
        
    'Prepara as Combos e a Mascara da Conta Contabil do Tab de Vendas
    lErro = CF("Inicializa_Campos_Obrigatorios", objUser)
    If lErro <> SUCESSO Then gError 126284
    
    lErro = Carrega_Usuarios
    If lErro <> SUCESSO Then gError 190461
    
    'Prepara as Combos  do Tab de Inscricao
    lErro = Inicializa_Tab_Inscricao()
    If lErro <> SUCESSO Then gError 207395
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207445
    
    Call CF2(Me, "CLI_Form_Load")
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 58084, 58085, 58086, 58087, 58505, 126284, 190461, 207395, 207445
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155569)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Function Inicializa_Tab_Identificacao() As Long

Dim lErro  As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Inicializa_Tab_Identificacao

'    'Lê Códigos e NomesReduzidos da tabela de Clientes e devolve na coleção
'    lErro = CF("LCod_Nomes_Le","Clientes", "Codigo", "NomeReduzido", STRING_CLIENTE_NOME_REDUZIDO, colCodigoNome)
'    If lErro <> SUCESSO Then Error 12262
'
'    'Preenche a ListBox ClientesList com os objetos da coleção
'    For Each objCodigoNome In colCodigoNome
'        ClientesList.AddItem objCodigoNome.sNome
'        ClientesList.ItemData(ClientesList.NewIndex) = objCodigoNome.lCodigo
'    Next

    'Lê cada código e descrição da tabela TiposDeCliente
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 12263

    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Tipo.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Tipo.ItemData(Tipo.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Set objGridCategoria = New AdmGrid

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then Error 28992

    'Carrega a combobox de Categoria Cliente
    lErro = Carrega_ComboCategoriaCliente()
    If lErro <> SUCESSO Then Error 28993

    Inicializa_Tab_Identificacao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Identificacao:

    Inicializa_Tab_Identificacao = Err

    Select Case Err
        
        Case 12262, 12263, 28992, 28993, 195759
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155570)
        
    End Select
        
    Exit Function
    
End Function

Private Function Inicializa_Tab_DadosFinanceiros() As Long

Dim lErro  As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Inicializa_Tab_DadosFinanceiros

    'Preenchimento ComboBox Tabela Preço
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela TabelasDePreco
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then Error 12264

    'Preenche a ComboBox TabelaPreco com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        TabelaPreco.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagto, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 12265

'    'Preenche Combo CondicaoPagto
'    Set colCodigoDescricao = New AdmColCodigoNome
'
'    'Lê cada codigo e descricao reduzida da tabela CondicoesPagto
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 12265
'
'    'Preenche a ComboBox CondicaoPagto com os objetos da colecao colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
'        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo
'    Next

    'Preenche Combo Mensagem
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela Mensagem
    lErro = CF("Cod_Nomes_Le", "Mensagens", "Codigo", "Descricao", STRING_NFISCAL_MENSAGEM, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 12266

    'Preenche a ComboBox Mensagem com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Mensagem.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Mensagem.ItemData(Mensagem.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    'Essa função só faz sentido para versões customizadas. Na versão 2 é uma função vazia
    lErro = CF("Inicializa_Tab_DadosFinanceiros_Usu", objUserControl)
    If lErro <> SUCESSO Then gError 127096
   
    Inicializa_Tab_DadosFinanceiros = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_DadosFinanceiros:
    
    Inicializa_Tab_DadosFinanceiros = Err
    
    Select Case Err
        
        Case 12264, 12265, 12266, 127096
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155571)
        
    End Select
        
    Exit Function
        
End Function

Private Function Inicializa_Tab_Vendas() As Long

Dim lErro  As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim sMascaraConta As String
Dim objCobrador As ClassCobrador
Dim ColCobrador As New Collection
Dim colPadroesCobranca As New Collection
Dim objPadraoCobranca As New ClassPadraoCobranca

On Error GoTo Erro_Inicializa_Tab_Vendas

    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 12271

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Regiao.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 12282
    
        ContaContabil.Mask = sMascaraConta

    Else
       
       'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 12282
    
        ContaContabil.Mask = sMascaraConta
       
        'Desabilita Conta Contabil
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
    End If
    

    'Preenche Combo Cobrador
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 12272

    'Preenche a ComboBox Cobrador com os objetos da colecao colCodigoDescricao
    For Each objCobrador In ColCobrador
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If
    Next

    'Preenche Combo Transportadora
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê códigos e nomes reduzidos da tabela Transportadoras
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 12273

    'Preenche as ComboBox Transportadora e TranspRedespacho com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Transportadora.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoDescricao.iCodigo
        TranspRedespacho.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TranspRedespacho.ItemData(TranspRedespacho.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    'Lê todos os Padroes de Cobranca da tabela PadroesCobranca
    lErro = CF("PadroesCobranca_Le_Todos", colPadroesCobranca)
    If lErro <> SUCESSO Then Error 58777
    
    For Each objPadraoCobranca In colPadroesCobranca
        
        'Verifica se Padrao de Cobranca está ativo
        If objPadraoCobranca.iInativo <> Inativo Then
            
            PadraoCobranca.AddItem CStr(objPadraoCobranca.iCodigo) & SEPARADOR & objPadraoCobranca.sDescricao
            PadraoCobranca.ItemData(PadraoCobranca.NewIndex) = objPadraoCobranca.iCodigo

        End If

    Next
    
    Inicializa_Tab_Vendas = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Vendas:
    
    Inicializa_Tab_Vendas = Err
    
    Select Case Err
    
        Case 12271, 12272, 12273, 12282, 58777
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155573)
        
    End Select
        
    Exit Function
            
End Function

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

Public Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub InscricaoMunicipal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub InscricaoSuframa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LimiteCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LimiteCredito_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sLimiteCredito As String

On Error GoTo Erro_LimiteCredito_Validate

    sLimiteCredito = LimiteCredito.Text

    'Verifica se foi preenchida a ComboBox LimiteCredito
    If Len(Trim(sLimiteCredito)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(sLimiteCredito)
    If lErro <> SUCESSO Then Error 12312

    'Formata
    LimiteCredito.Text = Format(sLimiteCredito, "Fixed")

    Call Calcula_SaldodeCredito

    Exit Sub

Erro_LimiteCredito_Validate:

    Cancel = True


    Select Case Err

        Case 12312

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155575)

    End Select


    Exit Sub

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

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155576)

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

Public Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 25001

    End If

    'Preenche os labels cliente existentes nos frames
    For iIndice = 0 To ClienteLabel.UBound

        ClienteLabel(iIndice).Caption = Trim(NomeReduzido.Text)

    Next

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True


    Select Case Err

        Case 25001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155577)

    End Select

End Sub

'Private Sub objEventoClienteFilial_evSelecao(obj1 As Object)
'
'Dim objFilialCliente As ClassFilialCliente
'Dim bCancel As Boolean
'
'    Set objFilialCliente = obj1
'
'    Codigo.Text = objFilialCliente.lCodCliente
'
'    Call Codigo_Validade(bCancel)
'
'    Me.Show
'
'    Exit Sub
'
'End Sub

Private Sub objEventoCobrador_evSelecao(obj1 As Object)

Dim objCobrador As ClassCobrador
Dim bCancel As Boolean

    Set objCobrador = obj1

    'Preenche campo Agente Cobrador
    Cobrador.Text = objCobrador.iCodigo

    'Chama a rotina de validate
    Call Cobrador_Validate(bCancel)

    Me.Show

    Exit Sub

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

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora
Dim bCancel As Boolean

    Set objTransportadora = obj1

    'Preenche campo Transportadora
    Transportadora.Text = objTransportadora.iCodigo

    'Chama a rotina de validate
    Call Transportadora_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_CLIENTES_ID
                
            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_CLIENTES_DADOS_FIN
            
            Case TAB_Inscricoes
                Parent.HelpContextID = IDH_CLIENTES_INSCRICOES
            
            Case TAB_Enderecos
                Parent.HelpContextID = IDH_CLIENTES_ENDERECOS
            
            Case TAB_Vendas
                Parent.HelpContextID = IDH_CLIENTES_VENDAS
            
            Case TAB_Estatisticas
                Parent.HelpContextID = IDH_CLIENTES_ESTATISTICAS
                        
        End Select
        
    End If

End Sub

Public Sub OpcaoEndereco_Click(Index As Integer)

    'Se frame selecionado for igual ao atual então não faz nada .
    If iFrame2Atual <> Index Then
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
        iFrame2Atual = Index
    End If

End Sub

Public Sub PadraoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PadraoCobranca_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PadraoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objPadraoCobranca As New ClassPadraoCobranca
Dim iCodigo As Integer

On Error GoTo Erro_PadraoCobranca_Validate

    'verifica se foi preenchido o campo PadraoCobranca
    If Len(Trim(PadraoCobranca.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox PadraoCobranca
    If PadraoCobranca.Text = PadraoCobranca.List(PadraoCobranca.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(PadraoCobranca, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 58778

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objPadraoCobranca.iCodigo = iCodigo

        'Tenta ler Padrao Cobranca com esse código no BD
        lErro = CF("PadraoCobranca_Le", objPadraoCobranca)
        If lErro <> SUCESSO And lErro <> 19298 Then Error 58779

        If lErro = 19298 Then Error 58780 'Não encontrou Padrao Cobranca no BD

        'Encontrou Padrao Cobranca no BD, coloca no Text da Combo
        PadraoCobranca.Text = CStr(objPadraoCobranca.iCodigo) & SEPARADOR & objPadraoCobranca.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 58781

    Exit Sub

Erro_PadraoCobranca_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 58778, 58779

    Case 58780  'Não encontrou Padrao Cobranca no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PADRAO_COBRANCA")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("PadroesCobranca", objPadraoCobranca)

        End If

    Case 58781
        lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", Err, PadraoCobranca.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155579)

    End Select

    Exit Sub

End Sub

Public Sub PadraoCobrancaLabel_Click()
'chama browse de tipos de padrao de cobranca

Dim objPadraoCobranca As New ClassPadraoCobranca
Dim colSelecao As New Collection

    If PadraoCobranca.ListIndex <> -1 Then objPadraoCobranca.iCodigo = Codigo_Extrai(PadraoCobranca.Text)

    Call Chama_Tela("PadraoCobrancaLista", colSelecao, objPadraoCobranca, objEventoPadraoCobranca)

End Sub

Private Sub objEventoPadraoCobranca_evSelecao(obj1 As Object)
'retorno do browse de tipos de padrao de cobranca

Dim objPadraoCobranca As ClassPadraoCobranca
Dim lErro As Long
Dim bCancel As Boolean

    Set objPadraoCobranca = obj1

    If objPadraoCobranca.iCodigo = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objPadraoCobranca.iCodigo)
        Call PadraoCobranca_Validate(bCancel)
    End If
    
    iAlterado = 0

    Me.Show

    Exit Sub

End Sub

Public Sub RazaoSocial_Change()

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

    'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
    If Regiao.Text = Regiao.List(Regiao.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19235

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 19236
        
        'Não encontrou Regiao Venda BD
        If lErro <> SUCESSO Then Error 19237
        
        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19238

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19235, 19236

    Case 19237  'Não encontrou RegiaoVenda no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

        If vbMsgRes = vbYes Then
            'Chama a tela RegiaoVenda
            Call Chama_Tela("RegiaoVenda", objRegiaoVenda)

        End If

    Case 19238
        lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", Err, Regiao.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155581)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabelaPreco_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 6732

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then Error 6737
        
        If lErro <> SUCESSO Then Error 6738 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 6739

    Exit Sub

Erro_TabelaPreco_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 6732, 6737

    Case 6738  'Não encontrou Tabela de Preço no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then

            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            
        End If

    Case 6739
        lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", Err, TabelaPreco.Text)

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155582)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Click()

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Tipo_Click

    'Configura flag de Alteração
    iAlterado = REGISTRO_ALTERADO

    'Se não foi selecionado Tipo, sai
    If Tipo.ListIndex = -1 Then Exit Sub

    objTipoCliente.iCodigo = Tipo.ItemData(Tipo.ListIndex)

    'Lê TipoCliente no BD
    lErro = CF("TipoDeCliente_Le", objTipoCliente)
    If lErro <> SUCESSO And lErro <> 28943 Then Error 33005

    'Tipo não está cadastrado
    If lErro = 28943 Then Error 33006

    lErro = Exibe_Dados_TipoCliente(objTipoCliente)
    If lErro <> SUCESSO Then Error 33007

    Exit Sub

Erro_Tipo_Click:

    Select Case Err

        Case 33005, 33007
            Tipo.SetFocus

        Case 33006  'Tipo não está cadastrado

            'Envia aviso perguntando se deseja cadastrar novo Tipo de Cliente
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoCliente", objTipoCliente)

            Else
                Tipo.SetFocus

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155583)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Tipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19219

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoCliente.iCodigo = iCodigo

        'Tenta ler TipoCliente com esse código no BD
        lErro = CF("TipoDeCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 28943 Then Error 19220

        'Não encontrou Tipo Cliente no BD
        If lErro = 28943 Then Error 33000

        'Exibe dados de TipoCliente na tela
        Tipo.Text = CStr(iCodigo) & SEPARADOR & objTipoCliente.sDescricao

        lErro = Exibe_Dados_TipoCliente(objTipoCliente)
        If lErro <> SUCESSO Then Error 6841

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19222

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case Err

        Case 6841, 19219, 19220  'Já tratado na rotina chamada

        Case 33000 'Não encontrou Tipo Cliente no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a tela de TiposDeClientes
                Call Chama_Tela("TipoCliente", objTipoCliente)

            End If

        Case 19222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155584)

    End Select

    Exit Sub

End Sub

Public Sub TipoClienteLabel_Click()

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE TIPO_CLIENTE :

Dim objTipoCliente As New ClassTipoCliente
Dim colSelecao As New Collection
    
    'Se o tipo estiver preenchido então passa para objTipoCliente
    If Len(Trim(Tipo.Text)) > 0 Then objTipoCliente.iCodigo = Codigo_Extrai(Tipo.Text)
    
    'Chama a tela que lista os Clientes
    Call Chama_Tela("TipoClienteLista", colSelecao, objTipoCliente, objEventoTipoCliente)

End Sub

Public Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

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
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19246

    Exit Sub

Erro_Transportadora_Validate:
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155585)

    End Select

    Exit Sub

End Sub

Public Sub TransportadoraLabel_Click()
'BROWSE TRANSPORTADORA :

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(Transportadora.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = 1

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
            If lErro <> SUCESSO Then gError 25024
            
            lErro = CF("Vendedor_Comissao_Customiza", objVendedor, dPercComissao)
            If lErro <> SUCESSO Then gError 126953
            
            If Len(Trim(ComissaoVendas.Text)) = 0 And dPercComissao <> 0 Then ComissaoVendas.Text = CStr(100 * dPercComissao)

            If objVendedor.iAtivo = DESMARCADO Then gError 140220 'Inserido por Wagner

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 25024, 126953   'Tratado na rotina chamada
        
        Case 140220
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155586)
    
    End Select

End Sub

Function Limpa_Tela_Clientes() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Clientes

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa as TextBox e as MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa os textos das Combos
    Tipo.Text = ""
    TabelaPreco.Text = ""
    CondicaoPagto.Text = ""
    Mensagem.Text = ""
    PadraoCobranca.Text = ""
    
    UsuRespCallCenter.Text = ""
    ComboCobrador.Text = ""
    
    Ativo.Value = MARCADO
    
    Bloqueado.Value = DESMARCADO 'Inserido por Wagner

    Regiao.Text = ""
    Cobrador.Text = ""
    TipoFrete.ListIndex = -1
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False

    ComboCobrador.ListIndex = -1
    UsuRespCallCenter.ListIndex = -1

    'Limpa os Labels cliente
    For iIndice = 0 To ClienteLabel.UBound
        ClienteLabel(iIndice).Caption = ""
    Next
    
    IEIsento.Value = vbChecked
    IENaoContrib.Value = vbChecked

    'Escolhe a opção Endereco Principal
    OpcaoEndereco(0).Value = True
    OpcaoEndereco(1).Value = False
    OpcaoEndereco(2).Value = False

    'Torna visível a Frame do Endereço Principal
    Frame2(iFrame2Atual).Visible = False
    Frame2(0).Visible = True
    iFrame2Atual = 0

    Call Grid_Limpa(objGridCategoria)
    
    If gobjFAT.iExibeRotasCliente = MARCADO Then
        Call Grid_Limpa(objGridRotas)
        RotaTodasFiliais.Value = vbChecked
        RotaFilialLista.ListIndex = -1
        RotaFilialLista.Enabled = False
    End If

    'limpa a estatítistica
    SaldoTitulos.Caption = ""
    SaldoDuplicatas.Caption = ""
    SaldoPedidosLiberados.Caption = ""
    NumeroCompras.Caption = ""
    MediaCompra.Caption = ""
    ValorAcumuladoCompras.Caption = ""
    
    TotalCR.Caption = ""
    TotalCREmCartorio.Caption = ""
    TotalCRComProtesto.Caption = ""
    PercCREmAberto.Caption = ""
    PercCREmCartorio.Caption = ""
    PercCRComProtesto.Caption = ""


    DataPrimeiraCompra.Caption = ""
    
    DataUltimaCompra.Caption = ""
    
    MediaAtraso.Caption = ""
    MaiorAtraso.Caption = ""
    SaldoAtrasados.Caption = ""
    ValorPagtosAtraso.Caption = ""
    
    NumChequesDevolvidos.Caption = ""

    DataUltChequeDevolvido.Caption = ""
    
    'Limpa o código de identificação
    Codigo.Text = ""
    
    ComboCobrador.ListIndex = -1
    
    Set gcolRotas = New Collection
    
    Call gobjTabEnd.Limpa_Tela
    
    Call CF2(Me, "CLI_Limpa_Tela_Cliente")
    
    Set gobjAnotacao = Nothing

    Limpa_Tela_Clientes = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_Clientes:

    Limpa_Tela_Clientes = Err
    
    Select Case Err
        
        Case 58584
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155587)

    End Select
    
    Exit Function
        
End Function

Function CLI_Trata_Parametros(ByVal objCT As Object) As Long
    Set objCT.gobjTabEnd = New ClassTabEndereco
End Function

Function Trata_Parametros(Optional objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Trata_Parametros

'    Call CF2(Me, "CLI_Trata_Parametros")

    'Se houver Cliente passado como parâmetro, exibe seus dados
    If Not (objCliente Is Nothing) Then

        'Se Codigo é positivo
        If objCliente.lCodigo > 0 Then

            'Lê Cliente no BD a partir do código
            lErro = CF("Cliente_Le_Estendida", objCliente, objClienteEstatistica)
            If lErro <> SUCESSO And lErro <> 52545 Then Error 12283

            'Se não encontrou o Cliente no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o código
                lErro = Limpa_Tela_Clientes()
                If lErro <> SUCESSO Then Error 58585
                
                Codigo.Text = CStr(objCliente.lCodigo)

            Else  'Encontrou Cliente no BD

                'Exibe os dados do Cliente
                lErro = Exibe_Dados_Cliente(objCliente, objClienteEstatistica)
                If lErro <> SUCESSO Then Error 12500

            End If

        'se Nome Reduzido está preenchido
        ElseIf Len(Trim(objCliente.sNomeReduzido)) > 0 Then

            'Lê Cliente no BD a partir do Nome Reduzido
            lErro = CF("Cliente_Le_NomeRed_Estendida", objCliente, objClienteEstatistica)
            If lErro <> SUCESSO And lErro <> 52693 Then Error 6923

            'Se não encontrou o Cliente no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o NomeReduzido
                lErro = Limpa_Tela_Clientes()
                If lErro <> SUCESSO Then Error 58586
                
                NomeReduzido.Text = CStr(objCliente.sNomeReduzido)

            Else  'Encontrou Cliente no BD

                'Exibe os dados do Cliente
                lErro = Exibe_Dados_Cliente(objCliente, objClienteEstatistica)
                If lErro <> SUCESSO Then Error 6924

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 6923, 6924, 12283, 12500, 58585, 58586  'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155588)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Exibe_Dados_Cliente(objCliente As ClassCliente, objClienteEstatistica As ClassFilialClienteEst) As Long
'Exibe os dados de Cliente na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_Cliente

    gbCarregandoTela = True
    
    'Limpa a Tela e exibe apenas o código
    lErro = Limpa_Tela_Clientes()
    If lErro <> SUCESSO Then gError 195841

    '#####################################
    'Inserido por Wagner
    'Lê os dados do cliente a ser excluido
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 12431

    'Verifica se cliente não está cadastrado
    If lErro = 12293 Then gError 12432
    '#####################################

    'Preenche os campos cliente existentes nos frames
    For iIndice = 0 To ClienteLabel.UBound
        ClienteLabel(iIndice).Caption = objCliente.sNomeReduzido
    Next

    'TAB IDENTIFICACAO :
    lErro = Exibe_Dados_Cliente_Identificacao(objCliente)
    If lErro <> SUCESSO Then gError 58095
    
    'TAB DADOS FINANCEIROS :
    Call Exibe_Dados_Cliente_DadosFinanceiros(objCliente)
    
    'TAB INSCRICOES :
    Call Exibe_Dados_Cliente_Inscricoes(objCliente)
    
    'TAB ENDERECOS :
    lErro = Exibe_Dados_Cliente_Enderecos(objCliente)
    If lErro <> SUCESSO Then gError 58096
    
    'TAB VENDAS:
    lErro = Exibe_Dados_Cliente_Vendas(objCliente)
    If lErro <> SUCESSO Then gError 58097
    
    'TAB ESTATISTICAS:
    Call Exibe_Dados_Cliente_Estatisticas(objCliente, objClienteEstatistica)
    
    'Faixa de Recebimento
    If objCliente.iIgnoraRecebPadrao = DESMARCADO Then
        IgnoraRecebPadrao.Value = vbUnchecked
    Else
        IgnoraRecebPadrao.Value = vbChecked
    End If
    
    'Faixa de Recebimento
    If objCliente.iTemFaixaReceb = DESMARCADO Then
        NaoTemFaixaReceb.Value = vbUnchecked
    Else
        NaoTemFaixaReceb.Value = vbChecked
    End If
    
    PercentMenosReceb.Text = (objCliente.dPercentMenosReceb) * 100
    PercentMaisReceb.Text = (objCliente.dPercentMaisReceb) * 100

    RecebForaFaixa(objCliente.iRecebForaFaixa).Value = True
    
    lErro = CF2(Me, "CLI_Traz_Cliente_Tela", objCliente)
    If lErro <> SUCESSO Then gError 192966
    
    If gobjFAT.iExibeRotasCliente = MARCADO Then
        lErro = Traz_Rotas_Tela(objCliente.lCodigo)
        If lErro <> SUCESSO Then gError 192966
    End If
    
    iAlterado = 0

    gbCarregandoTela = False
    
    Exibe_Dados_Cliente = SUCESSO

    Exit Function

Erro_Exibe_Dados_Cliente:
    
    gbCarregandoTela = False
    
    Exibe_Dados_Cliente = gErr

    Select Case gErr
        
        '######################
        'Inserido por Wagner
        Case 12431, 192966, 195841
        
        Case 12432
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
        '######################
        
        Case 58095, 58096, 58097
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155589)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Cliente_Identificacao(objCliente As ClassCliente) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoCliente As New ClassTipoCliente
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim iCodigo As Integer

On Error GoTo Erro_Exibe_Dados_Cliente_Identificacao
    
    Ativo.Value = objCliente.iAtivo
    Codigo.Text = CStr(objCliente.lCodigo)
    RazaoSocial.Text = objCliente.sRazaoSocial
    NomeReduzido.Text = objCliente.sNomeReduzido

    If objCliente.iTipo = 0 Then
        Tipo.Text = ""
    Else
        'Coloca Tipo no Text
        Tipo.Text = CStr(objCliente.iTipo)

        'Tenta selecionar
        lErro = Combo_Seleciona(Tipo, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And 6731 Then Error 33931

        If lErro <> SUCESSO Then

            objTipoCliente.iCodigo = objCliente.iTipo

            'Tenta ler TipoCliente com esse código no BD
            lErro = CF("TipoDeCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 28943 Then Error 12517

            If lErro = 28943 Then Error 6843 'Não encontrou Tipo Cliente no BD

            'Insere na Combo
            Tipo.AddItem CStr(objTipoCliente.iCodigo) & SEPARADOR & objTipoCliente.sDescricao
            Tipo.ItemData(Tipo.NewIndex) = objTipoCliente.iCodigo

            'Seleciona na Combo
            lErro = Combo_Seleciona(Tipo, iCodigo)
            If lErro <> SUCESSO Then Error 33747

        End If

    End If

    'CATEGORIAS

    'Lê as categoria deste cliente
    objFilialCliente.lCodCliente = objCliente.lCodigo
    objFilialCliente.iCodFilial = FILIAL_MATRIZ
    
    lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
    If lErro <> SUCESSO Then Error 33010

    'Limpa o grid de categorias
    Call Grid_Limpa(objGridCategoria)

    'Carrega o Grid Categoria
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = colCategoriaItem.Item(iIndice).sCategoria
            GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = colCategoriaItem.Item(iIndice).sItem
        Next
        objGridCategoria.iLinhasExistentes = colCategoriaItem.Count
    End If

    Observacao.Text = objCliente.sObservacao

    Exibe_Dados_Cliente_Identificacao = Err
    
    Exit Function
    
Erro_Exibe_Dados_Cliente_Identificacao:

    Exibe_Dados_Cliente_Identificacao = SUCESSO
    
    Select Case Err
    
        Case 6843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objTipoCliente.iCodigo)

        Case 12517, 33010, 33931

        Case 33747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_SELECIONADO", Err, objTipoCliente.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155590)

    End Select

    Exit Function

End Function

Private Sub Exibe_Dados_Cliente_DadosFinanceiros(objCliente As ClassCliente)
'Exibe os dados financeiros do Cliente

Dim bCancel As Boolean
    
    LimiteCredito.Text = CStr(objCliente.dLimiteCredito)

    '#######################################################
    'Inserido por Wagner
    If objCliente.iBloqueado = MARCADO Then
        Bloqueado.Value = vbChecked
    Else
        Bloqueado.Value = vbUnchecked
    End If
    '#######################################################

    If objCliente.iTabelaPreco = 0 Then
        TabelaPreco.Text = ""
    Else
        TabelaPreco.Text = CStr(objCliente.iTabelaPreco)
        Call TabelaPreco_Validate(bCancel)
    End If

    If objCliente.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objCliente.iCondicaoPagto)
        Call CondicaoPagto_Validate(bCancel)
    End If

    If objCliente.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objCliente.iCodMensagem)
        Call Mensagem_Validate(bCancel)
    End If

    Desconto.Text = CStr(100 * objCliente.dDesconto)

End Sub

Private Sub Exibe_Dados_Cliente_Inscricoes(objCliente As ClassCliente)
'Exibe as inscrições do Cliente
    
    Dim iIndice As Integer
    
    RG.Text = objCliente.sRG
    CGC.Text = objCliente.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objCliente.sInscricaoEstadual
    Call Trata_IE
    If objCliente.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    If objCliente.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If
        
    InscricaoMunicipal.Text = objCliente.sInscricaoMunicipal
    Observacao2.Text = objCliente.sObservacao2
    
    'Busca no ItemData da Combo de Regime Tributario o código
    For iIndice = 0 To RegimeTributario.ListCount - 1
        If RegimeTributario.ItemData(iIndice) = objCliente.iRegimeTributario Then
            RegimeTributario.ListIndex = iIndice
            Exit For
        End If
    Next

    
    InscricaoSuframa.PromptInclude = False
    InscricaoSuframa.Text = objCliente.sInscricaoSuframa
    InscricaoSuframa.PromptInclude = True
    
End Sub
    
Private Function Exibe_Dados_Cliente_Enderecos(objCliente As ClassCliente) As Long
'Exibe os Endereços do Cliente

Dim lErro As Long
Dim colEnderecos As New colEndereco

On Error GoTo Erro_Exibe_Dados_Cliente_Enderecos

    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
    If lErro <> SUCESSO Then gError 12304
    
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 196230

'    lErro = CF2(Me, "CLI_Exibe_Dados_Cliente_Enderecos", objCliente, colEnderecos)
'    If lErro <> SUCESSO Then gError 196230

    Exibe_Dados_Cliente_Enderecos = SUCESSO

    Exit Function

Erro_Exibe_Dados_Cliente_Enderecos:

    Exibe_Dados_Cliente_Enderecos = Err

    Select Case gErr

        Case 12304, 196230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155591)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Cliente_Vendas(objCliente As ClassCliente) As Long
'Exibe dados de Vendas do Cliente

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_Cliente_Vendas

    If objCliente.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objCliente.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If
    
    'Tipo de Frete
    If objCliente.iTipoFrete <> -1 Then
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objCliente.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoFrete.ListIndex = -1
    End If
    
    ContaContabil.PromptInclude = False
    If Len(Trim(objCliente.sContaContabil)) = 0 Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objCliente.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12494

        ContaContabil.Text = sContaEnxuta
    End If
    ContaContabil.PromptInclude = True

    If objCliente.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objCliente.dComissaoVendas)
    End If

    If objCliente.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objCliente.iRegiao)
        Call Regiao_Validate(bCancel)
    End If

    If objCliente.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objCliente.iCodCobrador)
        Call Cobrador_Validate(bCancel)
    End If

    If objCliente.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objCliente.iFreqVisitas)
    End If

    If (objCliente.dtDataUltVisita) = DATA_NULA Then
        DataUltVisita.PromptInclude = False
        DataUltVisita.Text = ""
        DataUltVisita.PromptInclude = True
    Else
        DataUltVisita.Text = Format(objCliente.dtDataUltVisita, "dd/mm/yy")
    End If

    If objCliente.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objCliente.iCodTransportadora)
        Call Transportadora_Validate(bCancel)
    End If

    If objCliente.iCodPadraoCobranca = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objCliente.iCodPadraoCobranca)
        Call PadraoCobranca_Validate(bCancel)
    End If

    Guia.Text = objCliente.sGuia
    
    If objCliente.iCodTranspRedesp = 0 Then
        TranspRedespacho.Text = ""
    Else
        TranspRedespacho.Text = CStr(objCliente.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If

    RedespachoCli.Value = objCliente.iDetPagFrete
    
    ComboCobrador.Text = objCliente.sUsuarioCobrador
    Call ComboCobrador_Validate(bSGECancelDummy)

    UsuRespCallCenter.Text = objCliente.sUsuRespCallCenter
    Call UsuRespCallCenter_Validate(bSGECancelDummy)

    Exibe_Dados_Cliente_Vendas = SUCESSO
    
    Exit Function
    
Erro_Exibe_Dados_Cliente_Vendas:

    Exibe_Dados_Cliente_Vendas = Err
    
    Select Case Err
            
        Case 12494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objCliente.sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155592)

    End Select

    Exit Function

End Function

Private Sub Exibe_Dados_Cliente_Estatisticas(objCliente As ClassCliente, objClienteEstatistica As ClassFilialClienteEst)
'Exibe dados de Estatísticas do Cliente

    SaldoTitulos.Caption = Format(objClienteEstatistica.dSaldoTitulos, "Standard")
    SaldoDuplicatas.Caption = Format(objClienteEstatistica.dSaldoDuplicatas, "Standard")
    SaldoPedidosLiberados.Caption = Format(objClienteEstatistica.dSaldoPedidosLiberados, "Standard")
    NumeroCompras.Caption = Format(objClienteEstatistica.lNumeroCompras, FORMATO_INTEIRO)
    MediaCompra.Caption = Format(objClienteEstatistica.dMediaCompra, "Standard")
    ValorAcumuladoCompras.Caption = Format(objClienteEstatistica.dValorAcumuladoCompras, "Standard")

    If objClienteEstatistica.dtDataPrimeiraCompra = DATA_NULA Then
        DataPrimeiraCompra.Caption = ""
    Else
        DataPrimeiraCompra.Caption = Format(objClienteEstatistica.dtDataPrimeiraCompra, "dd/mm/yyyy")
    End If

    If objClienteEstatistica.dtDataUltimaCompra = DATA_NULA Then
        DataUltimaCompra.Caption = ""
    Else
        DataUltimaCompra.Caption = Format(objClienteEstatistica.dtDataUltimaCompra, "dd/mm/yyyy")
    End If
    
    If objClienteEstatistica.lMediaAtraso > 1 Then
        MediaAtraso.Caption = CStr(objClienteEstatistica.lMediaAtraso) & " dias"
    Else
        MediaAtraso.Caption = CStr(objClienteEstatistica.lMediaAtraso) & " dia"
    End If

    If objClienteEstatistica.lMaiorAtraso > 1 Then
        MaiorAtraso.Caption = CStr(objClienteEstatistica.lMaiorAtraso) & " dias"
    Else
        MaiorAtraso.Caption = CStr(objClienteEstatistica.lMaiorAtraso) & " dia"
    End If

    SaldoAtrasados.Caption = Format(objClienteEstatistica.dSaldoAtrasados, "Standard")
    ValorPagtosAtraso.Caption = Format(objClienteEstatistica.dValPagtosAtraso, "Standard")
    
    NumChequesDevolvidos.Caption = Format(objClienteEstatistica.iNumChequesDevolvidos, FORMATO_INTEIRO)

    If objClienteEstatistica.dtDataUltChequeDevolvido = DATA_NULA Then
        DataUltChequeDevolvido.Caption = ""
    Else
        DataUltChequeDevolvido.Caption = Format(objClienteEstatistica.dtDataUltChequeDevolvido, "dd/mm/yyyy")
    End If
    
    TotalCR.Caption = Format(objClienteEstatistica.dValorTotalCR, "Standard")
    TotalCREmCartorio.Caption = Format(objClienteEstatistica.dValorTotalCREmCartorio, "Standard")
    TotalCRComProtesto.Caption = Format(objClienteEstatistica.dValorTotalCRComProtesto, "Standard")
    
    If objClienteEstatistica.dValorTotalCR > 0 Then
        PercCREmAberto.Caption = Format(objClienteEstatistica.dSaldoAtrasados / objClienteEstatistica.dValorTotalCR, "PERCENT")
        PercCREmCartorio.Caption = Format(objClienteEstatistica.dValorTotalCREmCartorio / objClienteEstatistica.dValorTotalCR, "PERCENT")
        PercCRComProtesto.Caption = Format(objClienteEstatistica.dValorTotalCRComProtesto / objClienteEstatistica.dValorTotalCR, "PERCENT")
    Else
        PercCREmAberto.Caption = Format(0, "PERCENT")
        PercCREmCartorio.Caption = Format(0, "PERCENT")
        PercCRComProtesto.Caption = Format(0, "PERCENT")
    End If
    
    Call Calcula_SaldodeCredito

End Sub

Function Exibe_Dados_TipoCliente(objTipoCliente As ClassTipoCliente) As Long
'Exibe dados do Tipo de Cliente na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim colCategoriaItem As New Collection
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_TipoCliente

    If gbCarregandoTela = False Then
    
        LimiteCredito.Text = CStr(objTipoCliente.dLimiteCredito)
    
        If objTipoCliente.iTabelaPreco = 0 Then
            TabelaPreco.Text = ""
        Else
            TabelaPreco.Text = CStr(objTipoCliente.iTabelaPreco)
            Call TabelaPreco_Validate(bCancel)
        End If
    
        If objTipoCliente.iCondicaoPagto = 0 Then
            CondicaoPagto.Text = ""
        Else
            CondicaoPagto.Text = CStr(objTipoCliente.iCondicaoPagto)
            Call CondicaoPagto_Validate(bCancel)
        End If
    
        If objTipoCliente.iCodMensagem = 0 Then
            Mensagem.Text = ""
        Else
            Mensagem.Text = CStr(objTipoCliente.iCodMensagem)
            Call Mensagem_Validate(bCancel)
        End If
    
        If objTipoCliente.dDesconto = 0# Then
            Desconto.Text = ""
        Else
            Desconto.Text = CStr(100 * objTipoCliente.dDesconto)
        End If
    
        ContaContabil.PromptInclude = False
    
        If objTipoCliente.sContaContabil = "" Then
            ContaContabil.Text = ""
        Else
            lErro = Mascara_RetornaContaEnxuta(objTipoCliente.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then Error 12491
    
            ContaContabil.Text = sContaEnxuta
        End If
    
        ContaContabil.PromptInclude = True
    
        If objTipoCliente.iVendedor = 0 Then
            Vendedor.Text = ""
        Else
            Vendedor.Text = CStr(objTipoCliente.iVendedor)
            Call Vendedor_Validate(bCancel)
        End If
    
        If objTipoCliente.dComissaoVendas = 0# Then
            ComissaoVendas.Text = ""
        Else
            ComissaoVendas.Text = CStr(100 * objTipoCliente.dComissaoVendas)
        End If
    
        If objTipoCliente.iFreqVisitas = 0 Then
            FreqVisitas.Text = ""
        Else
            FreqVisitas.Text = CStr(objTipoCliente.iFreqVisitas)
        End If
    
        If objTipoCliente.iCodTransportadora = 0 Then
            Transportadora.Text = ""
        Else
            Transportadora.Text = CStr(objTipoCliente.iCodTransportadora)
            Call Transportadora_Validate(bCancel)
        End If
    
        If objTipoCliente.iCodCobrador = 0 Then
            Cobrador.Text = ""
        Else
            Cobrador.Text = CStr(objTipoCliente.iCodCobrador)
            Call Cobrador_Validate(bCancel)
        End If
    
        If objTipoCliente.iRegiao = 0 Then
            Regiao.Text = ""
        Else
            Regiao.Text = CStr(objTipoCliente.iRegiao)
            Call Regiao_Validate(bCancel)
        End If
        
        If objTipoCliente.iPadraoCobranca = 0 Then
            PadraoCobranca.Text = ""
        Else
            PadraoCobranca.Text = CStr(objTipoCliente.iPadraoCobranca)
            Call PadraoCobranca_Validate(bCancel)
        End If
        'Lê as Categorias de Tipo de Cliente no BD
        lErro = CF("TipoDeClienteCategorias_Le", objTipoCliente, colCategoriaItem)
        If lErro <> SUCESSO Then Error 33011
    
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

    End If
    
    Exibe_Dados_TipoCliente = SUCESSO

    Exit Function

Erro_Exibe_Dados_TipoCliente:

    Exibe_Dados_TipoCliente = Err

    Select Case Err

        Case 33011

        Case 12491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTipoCliente.sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155593)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Cliente(objCliente As ClassCliente) As Long
'Lê os dados que estão na tela de Clientes e coloca em objCliente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objVendedor As New ClassVendedor
Dim objFilialCliCategoria As New ClassFilialCliCategoria
Dim objTabelaPreco As New ClassTabelaPreco
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objMensagem As New ClassMensagem
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objCobrador As New ClassCobrador
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Le_Dados_Cliente

    'IDENTIFICACAO :

    If Len(Trim(Codigo.Text)) > 0 Then objCliente.lCodigo = CLng(Codigo.Text)

    objCliente.sRazaoSocial = Trim(RazaoSocial.Text)
    objCliente.sNomeReduzido = Trim(NomeReduzido.Text)

    If Len(Trim(Tipo.Text)) > 0 Then objCliente.iTipo = Codigo_Extrai(Tipo.Text)

    objCliente.sObservacao = Trim(Observacao.Text)
    
    objCliente.iAtivo = Ativo.Value
    
    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialCliCategoria = New ClassFilialCliCategoria

            objFilialCliCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialCliCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objCliente.colCategoriaItem.Add objFilialCliCategoria

        End If

    Next

    'DADOS FINANCEIROS :

    '###################################
    'Inserido por Wagner
    If Bloqueado.Value = vbChecked Then
        objCliente.iBloqueado = MARCADO
    Else
        objCliente.iBloqueado = DESMARCADO
    End If
    '###################################

    If Len(Trim(LimiteCredito.Text)) > 0 Then objCliente.dLimiteCredito = CDbl(LimiteCredito.Text)
    If Len(Trim(Desconto.Text)) > 0 Then objCliente.dDesconto = CDbl(Desconto.Text) / 100

    If Len(Trim(TabelaPreco.Text)) > 0 Then objCliente.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    
    'Verifica se tabela de preço existe
    If objCliente.iTabelaPreco <> 0 Then
        objTabelaPreco.iCodigo = objCliente.iTabelaPreco
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 33013

        'Não encontrou a tabela de preço ==> erro
        If lErro = 28004 Then gError 33014

        If objTabelaPreco.iAtivo = DESMARCADO Then gError 213835

    End If

    objCliente.iCondicaoPagto = CondPagto_Extrai(CondicaoPagto)

    'Verifica se a condição de pagamento existe
    If objCliente.iCondicaoPagto <> 0 Then
        objCondicaoPagto.iCodigo = objCliente.iCondicaoPagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 33015

        'Não encontrou a condição de pagamento ==> erro
        If lErro = 19205 Then gError 33016

    End If

    If Len(Trim(Mensagem.Text)) > 0 Then objCliente.iCodMensagem = Codigo_Extrai(Mensagem.Text)

    'Verifica se mensagem existe
    If objCliente.iCodMensagem <> 0 Then
        objMensagem.iCodigo = objCliente.iCodMensagem
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 33017

        'Não encontrou a mensagem ==> erro
        If lErro = 19234 Then gError 33018

    End If

    'INSCRICOES :
    
    objCliente.sRG = Trim(RG.Text)
    objCliente.sCgc = Trim(CGC.Text)
    objCliente.sInscricaoEstadual = Trim(InscricaoEstadual.Text)
    objCliente.sInscricaoMunicipal = Trim(InscricaoMunicipal.Text)
    objCliente.sObservacao2 = Trim(Observacao2.Text)
    objCliente.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)
    objCliente.iRegimeTributario = RegimeTributario.ItemData(RegimeTributario.ListIndex)
    
    If IEIsento.Value = vbChecked Then
        objCliente.iIEIsento = MARCADO
    Else
        objCliente.iIEIsento = DESMARCADO
    End If
    If IENaoContrib.Value = vbChecked Then
        objCliente.iIENaoContrib = MARCADO
    Else
        objCliente.iIENaoContrib = DESMARCADO
    End If
    
    'VENDAS :

    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 33019

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 33020

        objCliente.iVendedor = objVendedor.iCodigo

    End If

    sConta = String(STRING_CONTA, 0)

    'Formata Conta contábil
    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 12364

    If iContaPreenchida = CONTA_VAZIA Then
        objCliente.sContaContabil = ""
    Else
        objCliente.sContaContabil = sConta
    End If

    If Len(Trim(ComissaoVendas.Text)) > 0 Then objCliente.dComissaoVendas = CDbl(ComissaoVendas.Text) / 100

    If Len(Trim(Regiao.Text)) > 0 Then objCliente.iRegiao = Codigo_Extrai(Regiao.Text)

    If Len(Trim(PadraoCobranca.Text)) > 0 Then objCliente.iCodPadraoCobranca = Codigo_Extrai(PadraoCobranca.Text)
    
    'Verifica se região venda existe
    If objCliente.iRegiao <> 0 Then
        objRegiaoVenda.iCodigo = objCliente.iRegiao
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 33093

        'Não encontrou a região venda ==> erro
        If lErro = 16137 Then gError 33094

    End If

    If Len(Trim(Cobrador.Text)) > 0 Then objCliente.iCodCobrador = Codigo_Extrai(Cobrador.Text)

    'Verifica se cobrador existe
    If objCliente.iCodCobrador <> 0 Then
        objCobrador.iCodigo = objCliente.iCodCobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 33021

        'Não encontrou o cobrador ==> erro
        If lErro = 19294 Then gError 33022

    End If

    If Len(Trim(FreqVisitas.Text)) > 0 Then objCliente.iFreqVisitas = CInt(FreqVisitas.Text)

    If Trim(DataUltVisita.Text) = "/  /" Then
        objCliente.dtDataUltVisita = DATA_NULA
    Else
        objCliente.dtDataUltVisita = CDate(DataUltVisita.Text)
    End If

    If TipoFrete.ListIndex <> -1 Then
        objCliente.iTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    Else
        objCliente.iTipoFrete = -1
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then objCliente.iCodTransportadora = Codigo_Extrai(Transportadora.Text)

    'Verifica se transportadora existe
    If objCliente.iCodTransportadora <> 0 Then
        objTransportadora.iCodigo = objCliente.iCodTransportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 33023

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 33024

    End If

    objCliente.sGuia = Trim(Guia.Text)
    
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objCliente.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)

    'Verifica se transportadora existe
    If objCliente.iCodTranspRedesp <> 0 Then
        objTransportadora.iCodigo = objCliente.iCodTranspRedesp
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 33023

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 33024

    End If

    objCliente.iDetPagFrete = RedespachoCli.Value
    
    objCliente.sUsuarioCobrador = ComboCobrador.Text
    objCliente.sUsuRespCallCenter = UsuRespCallCenter.Text
    
    lErro = CF2(Me, "CLI_Move_Cliente_Memoria", objCliente)
    If lErro <> SUCESSO Then gError 192967

    'Faixa de Recebimento
    If IgnoraRecebPadrao.Value = vbChecked Then
        objCliente.iIgnoraRecebPadrao = MARCADO
    ElseIf IgnoraRecebPadrao.Value = vbUnchecked Then
        objCliente.iIgnoraRecebPadrao = DESMARCADO
    End If
    
    'Faixa de Recebimento
    If NaoTemFaixaReceb.Value = vbChecked Then
        objCliente.iTemFaixaReceb = MARCADO
    ElseIf NaoTemFaixaReceb.Value = vbUnchecked Then
        objCliente.iTemFaixaReceb = DESMARCADO
    End If
    
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then
        objCliente.dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text) / 100
    Else
        objCliente.dPercentMenosReceb = 0
    End If

    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then
        objCliente.dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text) / 100
    Else
        objCliente.dPercentMaisReceb = 0
    End If

    If RecebForaFaixa(0).Value Then
        objCliente.iRecebForaFaixa = REJEITA_RECEBIMENTO
    End If

    If RecebForaFaixa(1).Value Then
        objCliente.iRecebForaFaixa = ACEITA_RECEBIMENTO
    End If

    Le_Dados_Cliente = SUCESSO

    Exit Function

Erro_Le_Dados_Cliente:

    Le_Dados_Cliente = gErr

    Select Case gErr
        
        'Erro tratado na rotina chamada
        Case 12364, 25010, 33013, 33015, 33017, 33019, 33021, 33023, 33093, 192967

        Case 33014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", gErr, objTabelaPreco.iCodigo)

        Case 33016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 33018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", gErr, objMensagem.iCodigo)

        Case 33020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 33022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)

        Case 33024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)

        Case 33094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case ERRO_SEM_MENSAGEM
        
        Case 213835
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_INATIVA", gErr, objTabelaPreco.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155594)

    End Select

    Exit Function

End Function
'
'Private Function Le_Dados_Enderecos(colEndereco As Collection) As Long
''Lê os dados relativos ao endereco(0 = principal , 1 = entrega , 2 = cobranca) do cliente e coloca em colEndereco
'
'Dim iIndice As Integer
'Dim objEndereco As ClassEndereco
'Dim iEstadoPreenchido As Integer
'
'On Error GoTo Erro_Le_Dados_Enderecos
'
'    'Verifica se tem algum estado Preenchido
'    For iIndice = 2 To 0 Step -1
'
'        If Len(Trim(Estado(iIndice).Text)) > 0 Then
'            iEstadoPreenchido = iIndice
'        End If
'
'    Next
'
'    For iIndice = 0 To 2
'
'        Set objEndereco = New ClassEndereco
'
'        objEndereco.sEndereco = Trim(Endereco(iIndice).Text)
'        objEndereco.sBairro = Trim(Bairro(iIndice).Text)
'        objEndereco.sCidade = Trim(Cidade(iIndice).Text)
'        objEndereco.sCEP = Trim(CEP(iIndice).Text)
'
'        'Se o Endereco não estiver Preenchido --> Seta o Estado que esta Preenchido em Algum dos Frames
'        If Len(Trim(Endereco(iIndice).Text)) > 0 Then
'            objEndereco.iCodigoPais = Codigo_Extrai(Pais(iIndice).Text)
'            objEndereco.sSiglaEstado = Trim(Estado(iIndice).Text)
'            If objEndereco.iCodigoPais = PAIS_BRASIL And Estado(iIndice).ListIndex = -1 Then Error 12325
'        Else
'            objEndereco.iCodigoPais = Codigo_Extrai(Pais(iEstadoPreenchido).Text)
'            If objEndereco.iCodigoPais = 0 Then Error 12326
'            objEndereco.sSiglaEstado = Trim(Estado(iEstadoPreenchido).Text)
'        End If
'
'        If objEndereco.iCodigoPais <> PAIS_BRASIL Then objEndereco.sSiglaEstado = ""
'
'        objEndereco.sTelefone1 = Trim(Telefone1(iIndice).Text)
'        objEndereco.sTelefone2 = Trim(Telefone2(iIndice).Text)
'        objEndereco.sFax = Trim(Fax(iIndice).Text)
'        objEndereco.sEmail = Trim(Email(iIndice).Text)
'        objEndereco.sContato = Trim(Contato(iIndice).Text)
'
'        colEndereco.Add objEndereco
'
'    Next
'
'    Le_Dados_Enderecos = SUCESSO
'
'    Exit Function
'
'Erro_Le_Dados_Enderecos:
'
'    Le_Dados_Enderecos = Err
'
'    Select Case Err
'
'        Case 12325
'            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Estado(iIndice).Text)
'
'        Case 12326
'            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_SELECIONADO", Err)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155595)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

'RETORNO TIPO_CLIENTE :
Private Sub objEventoTipoCliente_evSelecao(obj1 As Object)

Dim objTipoCliente As ClassTipoCliente
Dim lErro As Long
Dim bCancel As Boolean

    Set objTipoCliente = obj1

    'Preenche Text da Combo Tipo
    Tipo.Text = CStr(objTipoCliente.iCodigo)
    'Chama rotina de Validate
    Tipo_Validate (bCancel)

    iAlterado = 0

    Me.Show

    Exit Sub

End Sub

'RETORNO CONDICAO_PAGTO :
Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)

Dim objCondicaoPagto As ClassCondicaoPagto
Dim bCancel As Boolean

    Set objCondicaoPagto = obj1

    'Preenche Text da Combo CondicaoPagto
    CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo)
    'Chama rotina de Validate
    Call CondicaoPagto_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

'RETORNO VENDEDORES :
Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    iAlterado = 0

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

        'Mascara a conta
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 12851

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 12851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155596)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Clientes"

    'Lê os dados da Tela Clientes
    lErro = Le_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then Error 12584

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCliente.lCodigo, 0, "Codigo"
    colCampoValor.Add "RazaoSocial", objCliente.sRazaoSocial, STRING_CLIENTE_RAZAO_SOCIAL, "RazaoSocial"
    colCampoValor.Add "NomeReduzido", objCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Tipo", objCliente.iTipo, 0, "Tipo"
    colCampoValor.Add "Observacao", objCliente.sObservacao, STRING_CLIENTE_OBSERVACAO, "Observacao"
    colCampoValor.Add "LimiteCredito", objCliente.dLimiteCredito, 0, "LimiteCredito"
    colCampoValor.Add "CondicaoPagto", objCliente.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Desconto", objCliente.dDesconto, 0, "Desconto"
    colCampoValor.Add "CodMensagem", objCliente.iCodMensagem, 0, "CodMensagem"
    colCampoValor.Add "TabelaPreco", objCliente.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "CodPadraoCobranca", objCliente.iCodPadraoCobranca, 0, "CodPadraoCobranca"
    colCampoValor.Add "Ativo", objCliente.iAtivo, 0, "Ativo"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12584

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155597)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Tela_Preenche

    objCliente.lCodigo = colCampoValor.Item("Codigo").vValor

    If objCliente.lCodigo <> 0 Then

        'Carrega objCliente com os dados passados em colCampoValor
        objCliente.sRazaoSocial = colCampoValor.Item("RazaoSocial").vValor
        objCliente.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objCliente.iTipo = colCampoValor.Item("Tipo").vValor
        objCliente.sObservacao = colCampoValor.Item("Observacao").vValor
        objCliente.dLimiteCredito = colCampoValor.Item("LimiteCredito").vValor
        objCliente.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objCliente.dDesconto = colCampoValor.Item("Desconto").vValor
        objCliente.iCodMensagem = colCampoValor.Item("CodMensagem").vValor
        objCliente.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
        objCliente.iCodPadraoCobranca = colCampoValor.Item("CodPadraoCobranca").vValor
        objCliente.iAtivo = colCampoValor.Item("Ativo").vValor
        
        objClienteEstatistica.lCodCliente = objCliente.lCodigo
        objClienteEstatistica.iFilialEmpresa = giFilialEmpresa
                    
        lErro = CF("Cliente_Le_Estatistica", objClienteEstatistica)
        If lErro <> SUCESSO And lErro <> 52545 Then Error 52691
        
        'Inicializa objFilialCliente
        objFilialCliente.lCodCliente = objCliente.lCodigo
        objFilialCliente.iCodFilial = FILIAL_MATRIZ

        'Lê o restante dos dados do Cliente na tabela de Filiais
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then Error 19211

        'Carrega os restantes dos dados do Cliente contidos em objFilialCliente
        objCliente.sRG = objFilialCliente.sRG
        objCliente.sCgc = objFilialCliente.sCgc
        objCliente.sInscricaoEstadual = objFilialCliente.sInscricaoEstadual
        objCliente.sInscricaoMunicipal = objFilialCliente.sInscricaoMunicipal
        objCliente.sInscricaoSuframa = objFilialCliente.sInscricaoSuframa
        objCliente.lEndereco = objFilialCliente.lEndereco
        objCliente.lEnderecoEntrega = objFilialCliente.lEnderecoEntrega
        objCliente.lEnderecoCobranca = objFilialCliente.lEnderecoCobranca
        objCliente.iTipoFrete = objFilialCliente.iTipoFrete
        objCliente.iCodTransportadora = objFilialCliente.iCodTransportadora
        objCliente.iCodTranspRedesp = objFilialCliente.iCodTranspRedesp
        objCliente.iDetPagFrete = objFilialCliente.iDetPagFrete
        objCliente.sGuia = objFilialCliente.sGuia
        objCliente.sObservacao2 = objFilialCliente.sObservacao
        objCliente.sContaContabil = objFilialCliente.sContaContabil
        objCliente.iVendedor = objFilialCliente.iVendedor
        objCliente.dComissaoVendas = objFilialCliente.dComissaoVendas
        objCliente.iRegiao = objFilialCliente.iRegiao
        objCliente.iFreqVisitas = objFilialCliente.iFreqVisitas
        objCliente.dtDataUltVisita = objFilialCliente.dtDataUltVisita
        objCliente.iCodCobrador = objFilialCliente.iCodCobrador

        'Exibe o Cliente na Tela
        lErro = Exibe_Dados_Cliente(objCliente, objClienteEstatistica)
        If lErro <> SUCESSO Then Error 19213

'        'Desseleciona ListBox Clientes
'        ClientesList.ListIndex = -1

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 19211, 19213, 52691

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155598)

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
    
    Set gobjAnotacao = Nothing
    
    Call CF2(Me, "CLI_Form_UnLoad")
    
    Set gobjInfoUsu = Nothing
    
    'Descarrega variáveis globais tipo AdmEvento
    Set objEventoTipoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoCobrador = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoPadraoCobranca = Nothing
    Set objEventoCliente = Nothing
    'Set gobjTabEnd = Nothing
    Set objGridRotas = Nothing
    Set gcolRotas = Nothing

    'Descarrega variável global
    Set objGridCategoria = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

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
        If lErro <> SUCESSO Then Error 28995

    End If

    'Seleciona o Item na List da Combo
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

        Case 28995

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155599)

    End Select

End Function

Private Function Carrega_ComboCategoriaClienteItem(objCategoriaCliente As ClassCategoriaCliente) As Long
'Carrega a Combo CategoriaClienteItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaClienteItem As ClassCategoriaClienteItem

On Error GoTo Erro_Carrega_ComboCategoriaClienteItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then Error 28996

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

        Case 28996

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155600)

    End Select

    Exit Function

End Function

Public Sub Label1_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Codigo.Text)) > 0 Then objCliente.lCodigo = CDbl(Codigo.Text)

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub Label3_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim lErro As Long
Dim sNomeBrowse As String

On Error GoTo Erro_Label3_Click

    objCliente.sNomeReduzido = NomeReduzido.Text
    
    sNomeBrowse = "ClientesLista"
    
    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoCliente)

    Exit Sub

Erro_Label3_Click:

    Select Case gErr

        Case 71926

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155601)

    End Select

    Exit Sub


End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Executa o Validate
    Call Cliente_Traz_Tela(objCliente.lCodigo)

    Me.Show

    Exit Sub

End Sub

Public Sub Cliente_Traz_Tela(ByVal lCodigo As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Cliente_Traz_Tela

    'Guarda o valor do código do Cliente selecionado na ListBox ClientesList
    objCliente.lCodigo = lCodigo

    'Lê o Cliente no BD
    lErro = CF("Cliente_Le_Estendida", objCliente, objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52545 Then gError 71923

    'Se cliente não está cadastrado, erro
    If lErro = 12293 Then gError 71925

    'Exibe os dados do Cliente
    lErro = Exibe_Dados_Cliente(objCliente, objClienteEstatistica)
    If lErro <> SUCESSO Then gError 71924

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Cliente_Traz_Tela:

    Select Case gErr

        Case 71923, 71924

        Case 71925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155602)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaCliente() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos", colCategorias)
    If lErro <> SUCESSO Then Error 28997

    For Each objCategoriaCliente In colCategorias

        'Insere na combo CategoriaCliente
        ComboCategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next

    Carrega_ComboCategoriaCliente = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaCliente:

    Carrega_ComboCategoriaCliente = Err

    Select Case Err

        Case 28997

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155603)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridCategoria.Name

                Select Case GridCategoria.Col
        
                    Case GRID_CATEGORIA_COL
        
                        'Chama SaidaCelula de Categoria
                        lErro = Saida_Celula_Categoria(objGridInt)
                        If lErro <> SUCESSO Then gError 28998
        
                    Case GRID_VALOR_COL
        
                        'Chama SaidaCelula de Valor
                        lErro = Saida_Celula_Valor(objGridInt)
                        If lErro <> SUCESSO Then gError 28999
        
                End Select
                
            Case Else
            
                lErro = CF2(Me, "CLI_Saida_Celula", objGridInt)
                If lErro <> SUCESSO Then gError 195769
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 33025

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 28998, 28999, 195769

        Case 33025
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155604)

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

    'Testa se .Text da combo é igual ao TextMatrix "embaixo"
    iIndice = GridCategoria.Row
    
    'Se não for,
    If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) <> ComboCategoriaCliente.Text Then

        If Len(Trim(ComboCategoriaCliente.Text)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaCliente.sCategoria = ComboCategoriaCliente.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaCliente_Le", objCategoriaCliente)
            If lErro <> SUCESSO And lErro <> 28847 Then Error 33026

            'Categoria não está cadastrada
            If lErro = 28847 Then Error 33027

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        'Se já existe a categoria, erro
                        If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaCliente.Text Then Error 33028

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

                'Categoria vazia, limpa o Item correspondente
                GridCategoria.TextMatrix(iIndice1, GRID_VALOR_COL) = ""

            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33029

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = Err

    Select Case Err

        Case 33026
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33027 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaCliente"
                Call Chama_Tela("CategoriaCliente", objCategoriaCliente)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 33028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33029
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155605)

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
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then Error 55907

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaClienteItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaClienteItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaClienteItem.sItem = ComboCategoriaClienteItem.Text

            'Lê Item De Categoria De Cliente no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaClienteItem)
            If lErro <> SUCESSO And lErro <> 28991 Then Error 33030

            'Item da Categoria não está cadastrado
            If lErro = 28991 Then Error 33031

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33032

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 33030
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 33031 'Item da Categoria não está cadastrado

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

        Case 33032
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55907
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155606)

    End Select

    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CLIENTES_ID
    Set Form_Load_Ocx = Me
    Caption = "Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Clientes"
    
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
        
        If Me.ActiveControl Is Tipo Then
            Call TipoClienteLabel_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call Label1_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call Label3_Click
        ElseIf Me.ActiveControl Is CondicaoPagto Then
            Call CondicaoPagtoLabel_Click
        ElseIf Me.ActiveControl Is Mensagem Then
            Call MensagemNFLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is Cobrador Then
            Call AgenteCobradorLabel_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        End If
    
    End If

    Call CF2(Me, "CLI_UserControl_KeyDown", KeyCode, Shift)

End Sub

Public Sub CLI_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
'
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
'''
'''Public Property Get BotaoAnotar() As Object
'''     Set BotaoAnotar = objUserControl.Controls("BotaoAnotar")
'''End Property
'''
'''Public Sub BotaoAnotar_Click()
'''
'''    Call Chama_Tela("Anotacoes", gobjAnotacao)
'''
'''End Sub
'''
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155607)

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

Public Sub RedespachoCli_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Guia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_CLIENTE
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155610)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155611)
     
    End Select
     
    Exit Sub

End Sub

'####################################################
'Inserido por Wagner
Public Sub Bloqueado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub
'####################################################

Public Sub BotaoContatos_Click()

Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_BotaoContatos_Click
    
    objClienteContatos.lCliente = StrParaDbl(Codigo.Text)
    objClienteContatos.iFilialCliente = FILIAL_MATRIZ
    
    Call Chama_Tela("ClienteContatos", objClienteContatos)

    Exit Sub
    
Erro_BotaoContatos_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select


End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 190456

    For Each objUsuarios In colUsuarios
        ComboCobrador.AddItem objUsuarios.sCodUsuario
        UsuRespCallCenter.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 190456

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190457)

    End Select

    Exit Function

End Function

Public Sub ComboCobrador_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_ComboCobrador_Validate
    
    If Len(Trim(ComboCobrador.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = ComboCobrador.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 190458
        
        If lErro <> SUCESSO Then gError 190459
    
    End If
    
    Exit Sub
    
Erro_ComboCobrador_Validate:

    Cancel = True

    Select Case gErr
            
        Case 190458
        
        Case 190459 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190460)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub UsuRespCallCenter_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UsuRespCallCenter_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_UsuRespCallCenter_Validate
    
    If Len(Trim(UsuRespCallCenter.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = UsuRespCallCenter.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 190458
        
        If lErro <> SUCESSO Then gError 190459
        
    End If
    
    Exit Sub
    
Erro_UsuRespCallCenter_Validate:

    Cancel = True

    Select Case gErr
            
        Case 190458
        
        Case 190459 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190460)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function CLI_Form_Load(ByVal objCT As Object) As Long
    CLI_Form_Load = SUCESSO
End Function

Public Function CLI_Form_UnLoad(ByVal objCT As Object) As Long
    CLI_Form_UnLoad = SUCESSO
End Function

Public Function CLI_Grava_Registro(ByVal objCT As Object) As Long
    CLI_Grava_Registro = SUCESSO
End Function

Public Function CLI_Grava_Registro1(ByVal objCT As Object, ByVal objCliente As ClassCliente, colEnderecos As Collection) As Long
    CLI_Grava_Registro1 = SUCESSO
End Function

Public Function CLI_Limpa_Tela_Cliente(ByVal objCT As Object) As Long
    CLI_Limpa_Tela_Cliente = SUCESSO
End Function

Public Function CLI_Move_Cliente_Memoria(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long
    CLI_Move_Cliente_Memoria = SUCESSO
End Function

Public Function CLI_Traz_Cliente_Tela(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long
    CLI_Traz_Cliente_Tela = SUCESSO
End Function

Private Function Calcula_SaldodeCredito() As Long

Dim lErro As Long

On Error GoTo Erro_Calcula_SaldodeCredito

    SaldoLimitedeCredito.Caption = Format(StrParaDbl(LimiteCredito.Text), "STANDARD")
    
    SaldodeCredito.Caption = Format(StrParaDbl(SaldoLimitedeCredito.Caption) - StrParaDbl(SaldoTitulos.Caption) - StrParaDbl(SaldoPedidosLiberados.Caption), "STANDARD")

    Calcula_SaldodeCredito = SUCESSO

    Exit Function

Erro_Calcula_SaldodeCredito:

    Calcula_SaldodeCredito = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function CLI_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    CLI_Saida_Celula = SUCESSO
End Function

Public Function CLI_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    CLI_Rotina_Grid_Enable = SUCESSO
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    
    lErro = CF2(Me, "CLI_Rotina_Grid_Enable", iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 195787
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 195787

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195788)

    End Select

    Exit Sub

End Sub

Public Sub NaoTemFaixaReceb_Click()

Dim lErro As Long

On Error GoTo Erro_NaoTemFaixaReceb_Click

    'Verifica valor na checkbox
    If NaoTemFaixaReceb.Value = vbUnchecked Then

        'Habilita os controles
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True

    Else

        PercentMaisReceb.Text = "0"
        PercentMenosReceb.Text = "0"

        'Desabilita os controles
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If
        
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_NaoTemFaixaReceb_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154686)

    End Select

    Exit Sub

End Sub

Public Sub IgnoraRecebPadrao_Click()

Dim lErro As Long

On Error GoTo Erro_IgnoraRecebPadrao_Click

    'Verifica valor na checkbox
    If IgnoraRecebPadrao.Value = vbChecked Then

        'Habilita os controles
        NaoTemFaixaReceb.Enabled = True
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True

    Else

        PercentMaisReceb.Text = "0"
        PercentMenosReceb.Text = "0"
        NaoTemFaixaReceb.Value = vbUnchecked

        'Desabilita os controles
        NaoTemFaixaReceb.Enabled = False
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If
        
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_IgnoraRecebPadrao_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154686)

    End Select

    Exit Sub

End Sub

Public Sub PercentMaisReceb_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisReceb, iAlterado)
    
End Sub

Public Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_PercentMaisReceb_Validate

    'Verifica se PercentMaisRecebe foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then gError 49340

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca o valor no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMaisReceb_Validate:

    Cancel = True
    
    Select Case gErr

        Case 49340

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154688)

    End Select

    Exit Sub

End Sub

Public Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMenosReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMenosReceb, iAlterado)
    
End Sub

Public Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMenosReceb As Double

On Error GoTo Erro_PercentMenosReceb_Validate

    'Verifica se PercentMenosRecebe foi preenchido
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMenosReceb.Text)
        If lErro <> SUCESSO Then gError 49342

        dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text)

        'Verifica se porcentagem igual a 100%
        If dPercentMenosReceb = 100 Then gError 49343

        'Coloca o valor no formato fixed da tela
        PercentMenosReceb.Text = Format(dPercentMenosReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMenosReceb_Validate:

    Cancel = True

    Select Case gErr

        Case 49342

        Case 49343
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154690)

    End Select

    Exit Sub

End Sub

Public Sub RecebForaFaixa_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Inicializa_Grid_Rotas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Chave")
    objGridInt.colColuna.Add ("Rota")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (RotaFilial.Name)
    objGridInt.colCampo.Add (RotaChave.Name)
    objGridInt.colCampo.Add (RotaCodigo.Name)

    'Colunas do Grid
    iGrid_RotaFilial_Col = 1
    iGrid_RotaChave_Col = 2
    iGrid_RotaCodigo_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridRotas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridRotas.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Rotas = SUCESSO

    Exit Function

End Function

Public Sub GridRotas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRotas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRotas, iAlterado)
    End If

End Sub

Public Sub GridRotas_EnterCell()

    Call Grid_Entrada_Celula(objGridRotas, iAlterado)

End Sub

Public Sub GridRotas_GotFocus()

    Call Grid_Recebe_Foco(objGridRotas)

End Sub

Public Sub GridRotas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRotas)

End Sub

Public Sub GridRotas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRotas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRotas, iAlterado)
    End If

End Sub

Public Sub GridRotas_LeaveCell()

    Call Saida_Celula(objGridRotas)

End Sub

Public Sub GridRotas_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRotas)

End Sub

Public Sub GridRotas_RowColChange()

    Call Grid_RowColChange(objGridRotas)

End Sub

Public Sub GridRotas_Scroll()

    Call Grid_Scroll(objGridRotas)

End Sub

Private Function Traz_Rotas_Tela(ByVal lCliente As Long, Optional iFilial As Integer = 0) As Long

Dim lErro As Long
Dim colRotas As New Collection
Dim objRota As ClassRotas
Dim iLinha As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Traz_Rotas_Tela

    If lCliente <> 0 Then
    
        RotaTodasFiliais.Value = vbChecked
    
        objCliente.lCodigo = lCliente

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 205263

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", RotaFilialLista, colCodigoNome)

        lErro = CF("Cliente_Le_Rotas", lCliente, colRotas)
        If lErro <> SUCESSO Then gError 205263
                
        Set gcolRotas = colRotas
        
    End If
    
    Call Grid_Limpa(objGridRotas)
    iLinha = 0
    For Each objRota In gcolRotas
        If iFilial = 0 Or iFilial = objRota.iFilial Then
            iLinha = iLinha + 1
            GridRotas.TextMatrix(iLinha, iGrid_RotaFilial_Col) = CStr(objRota.iFilial) & SEPARADOR & objRota.sFilial
            GridRotas.TextMatrix(iLinha, iGrid_RotaChave_Col) = CStr(objRota.lChave) & SEPARADOR & objRota.sChave
            GridRotas.TextMatrix(iLinha, iGrid_RotaCodigo_Col) = objRota.sCodigo
        End If
    Next
    objGridRotas.iLinhasExistentes = iLinha

    Traz_Rotas_Tela = SUCESSO

    Exit Function

Erro_Traz_Rotas_Tela:

    Traz_Rotas_Tela = gErr

    Select Case gErr
    
        Case 205263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub BotaoNovaRota_Click()
    Call Chama_Tela("Rotas")
End Sub

Public Sub BotaoRotaEditar_Click()

Dim objRota As New ClassRotas

On Error GoTo Erro_BotaoRotaEditar_Click

    If GridRotas.Row = 0 Then gError 205207
    
    objRota.iFilialEmpresa = giFilialEmpresa
    objRota.sCodigo = GridRotas.TextMatrix(GridRotas.Row, iGrid_RotaCodigo_Col)
    objRota.lChave = LCodigo_Extrai(GridRotas.TextMatrix(GridRotas.Row, iGrid_RotaChave_Col))

    Call Chama_Tela("Rotas", objRota)
    
    Exit Sub
    
Erro_BotaoRotaEditar_Click:
    
    Select Case gErr
    
        Case 205207
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205209)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub RotaFilialLista_Click()
     Call Traz_Rotas_Tela(0, Codigo_Extrai(RotaFilialLista.Text))
End Sub

Public Sub RotaTodasFiliais_Click()
     If RotaTodasFiliais.Value = vbChecked Then
        RotaFilialLista.ListIndex = -1
        RotaFilialLista.Enabled = False
     Else
        If RotaFilialLista.ListCount > 0 Then RotaFilialLista.ListIndex = 0
        RotaFilialLista.Enabled = True
     End If
     Call RotaFilialLista_Click
End Sub

Public Sub BotaoTitRec_Click()

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(Codigo.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("TitRecBaixaLista", colSelecao, Nothing, Nothing, "Cliente = ?")
    
End Sub

Public Sub BotaoTitRecComProt_Click()

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(Codigo.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("TitRecBaixaLista", colSelecao, Nothing, Nothing, "Cliente = ? AND ComProtesto = 'Sim'")

End Sub

Public Sub BotaoTitRecEmCart_Click()

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(Codigo.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("TitRecBaixaLista", colSelecao, Nothing, Nothing, "Cliente = ? AND EmCartorio = 'Sim'")

End Sub

Public Sub BotaoTitRecPgAtrasado_Click()

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(Codigo.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("TitRecBaixaLista", colSelecao, Nothing, Nothing, "Cliente = ? AND Saldo = 0 AND Atraso > 0")

End Sub

Public Sub BotaoTitRecVenc_Click()

Dim colSelecao As New Collection
    
    colSelecao.Add StrParaLong(Codigo.Text)
    
    'Chama a tela que lista os Cobradores
    Call Chama_Tela("TitRecBaixaLista", colSelecao, Nothing, Nothing, "Cliente = ? AND Saldo > 0 AND Atraso > 0")

End Sub

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207448

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
        Case 207448

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207449)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207396)
        
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
