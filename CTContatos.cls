VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTContatos"
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

''''==================== ANOTACAO =====================
'Public gobjAnotacao As ClassAnotacoes
''''===================================================
Public gobjTabEnd As New ClassTabEndereco

Dim iFrameAtual As Integer
Dim iFrame2Atual As Integer
Public iAlterado As Integer
Dim iVendedorAlterado As Integer

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Private WithEvents objEventoTipoContato As AdmEvento
Attribute objEventoTipoContato.VB_VarHelpID = -1
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
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DadosFinanceiros = 2
Private Const TAB_Inscricoes = 3
Private Const TAB_Enderecos = 4
Private Const TAB_Vendas = 5

'--- inicio dos properties get dos controles da tela
Public Property Get Bloqueado() As Object
     Set Bloqueado = objUserControl.Controls("Bloqueado")
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

Public Property Get ContatoLabel() As Object
     Set ContatoLabel = objUserControl.Controls("ContatoLabel")
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

Public Property Get ComboCategoriaContato() As Object
     Set ComboCategoriaContato = objUserControl.Controls("ComboCategoriaContato")
End Property

Public Property Get ComboCategoriaContatoItem() As Object
     Set ComboCategoriaContatoItem = objUserControl.Controls("ComboCategoriaContatoItem")
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

Public Property Get TipoContatoLabel() As Object
     Set TipoContatoLabel = objUserControl.Controls("TipoContatoLabel")
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

Public Property Get TabEnd() As Object
     Set TabEnd = objUserControl.Controls("TabEnd")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do próximo Contato
    lErro = CF("Contato_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 180590

    'Exibe código na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 180590
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180591)
    
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
Dim objContato As New ClassContatos
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 180592

    objContato.lCodigo = StrParaLong(Codigo.Text)

    'Lê os dados do Contato a ser excluido
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180551 Then gError 180593

    'Verifica se Contato não está cadastrado
    If lErro <> SUCESSO Then gError 180594

    'Envia aviso perguntando se realmente deseja excluir Contato e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CONTATO", objContato.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Contato
        lErro = CF("Contato_Exclui", objContato)
        If lErro <> SUCESSO Then gError 180595

        'Limpa a Tela
        lErro = Limpa_Tela_Contatos()
        If lErro <> SUCESSO Then gError 180596

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 180592
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONTATO_NAO_PREENCHIDO", gErr)

        Case 180593, 180595, 180596

        Case 180594
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180597)

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

Public Sub Cidade_Change(Index As Integer)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180598

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 180599
        
        'Não encontrou Cobrador no BD
        If lErro <> SUCESSO Then gError 180600
        
        If giFilialEmpresa <> EMPRESA_TODA Then
            
            If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
                If objCobrador.iFilialEmpresa <> giFilialEmpresa Then gError 180601
            End If
            
        End If

        'Encontrou Cobrador no BD, coloca no Text da Combo
        Cobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180602

    Exit Sub

Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180598, 180599

        Case 180600  'Não encontrou Cobrador no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

            If vbMsgRes = vbYes Then
    
                 Call Chama_Tela("Cobradores", objCobrador)
    
            End If
    
        Case 180601
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_PERTENCE_FILIAL", gErr, objCobrador.iCodigo, giFilialEmpresa)
    
        Case 180602
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, Cobrador.Text)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180603)
    
    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Codigo Contato
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é um Long
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 180604

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 180604

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180605)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaContato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaContato_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaContato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Public Sub ComboCategoriaContato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaContato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaContato
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComboCategoriaContatoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaContatoItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaContatoItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaContatoItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaContatoItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaContatoItem_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180606)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaContatoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaContatoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaContatoItem
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180607

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 180608
        
        'Não encontrou CondicaoPagto no BD
        If lErro <> SUCESSO Then gError 180609

        'Encontrou CondicaoPagto no BD e não é de Recebimento
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 180610

        'Coloca no Text da Combo
        CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180611

    Exit Sub

Erro_CondicaoPagto_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180607, 180608

        Case 180609  'Não encontrou CondicaoPagto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If

        Case 180610
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", gErr, iCodigo)

        Case 180611
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180612)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection
    
    'Se a codicao estiver preenchida então passa para o objCondicaoPagamento
    If Len(Trim(CondicaoPagto.Text)) > 0 Then objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagto.Text)
    
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
    If lErro <> SUCESSO Then gError 180613
    
    'Se a conta estiver preenchida --> passa para o objPlanoConta
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta
    
    'Chama a lista de Contas
    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

    Case 180613

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180614)

    End Select

    Exit Sub

End Sub

Public Sub DataUltVisita_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataUltVisita, iAlterado)

End Sub

Public Sub Filiais_Click()

Dim lErro As Long
Dim objContato As New ClassContatos
Dim objFilialContato As New ClassFilialContato
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filiais_Click

    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 180615

    'Preenche objContato
    objContato.lCodigo = CLng(Codigo.Text)
    
    'Lê o Contato
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 180616

    'Se não achou o Contato
    If lErro <> SUCESSO Then

        'Envia aviso perguntando se deseja cadastrar novo Contato
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO")

        If vbMsgRes = vbYes Then
            
            'Grava o novo Contato
            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then gError 180617
            
            'Chama a Tela de Filiais de Contato
            objFilialContato.lCodContato = StrParaLong(Codigo.Text)
            objFilialContato.iCodFilial = 1 'p/começar exibindo a matriz
            
            Call Chama_Tela("FiliaisContatos", objFilialContato)
        
        End If
        
    Else
    
        'Chama a Tela de Filiais de Contato
        objFilialContato.lCodContato = StrParaLong(Codigo.Text)
        objFilialContato.iCodFilial = 1 'p/começar exibindo a matriz
        
        Call Chama_Tela("FiliaisContatos", objFilialContato)
    
    End If

    Exit Sub

Erro_Filiais_Click:

    Select Case gErr

        Case 180615
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONTATO_NAO_PREENCHIDO", gErr)

        Case 180616, 180617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180618)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Contato
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 180619

    'Limpa a Tela
    lErro = Limpa_Tela_Contatos()
    If lErro <> SUCESSO Then gError 180620
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 180619, 180620

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180621)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Contato necessários foram preenchidos
'Grava Contato no BD
'Atualiza ListBox de Contatos

Dim lErro As Long
Dim iIndice As Integer
Dim objContato As New ClassContatos
Dim colEndereco As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 180622

    'Verifica se foi preenchida a Razao Social
    If Len(Trim(RazaoSocial.Text)) = 0 Then gError 180623

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 180624

'    If giTipoVersao = VERSAO_FULL Then
        'Verifica se foi preenchido o Tipo
        If Len(Trim(Tipo.Text)) = 0 Then gError 180625
'    End If

'    'Verifica se foi preenchido o Estado dos Endereços
'    If Len(Trim(Endereco(0).Text)) <> 0 Then
'        If Len(Trim(Estado(0).Text)) = 0 Then gError 180626
'    End If
'
'    If Len(Trim(Endereco(1).Text)) <> 0 Then
'        If Len(Trim(Estado(1).Text)) = 0 Then gError 180627
'    End If
'
'    If Len(Trim(Endereco(2).Text)) <> 0 Then
'        If Len(Trim(Estado(2).Text)) = 0 Then gError 180628
'    End If

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
         If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 180629
    Next

'    'Lê os dados dos 3 três Enderecos e coloca em colEndereco
'    lErro = Le_Dados_Enderecos(colEndereco)
'    If lErro <> SUCESSO Then gError 180630
    
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEndereco)
    If lErro <> SUCESSO Then gError 180630


    'Lê os dados da Tela relacionados ao Contato
    lErro = Le_Dados_Contato(objContato)
    If lErro <> SUCESSO Then gError 180631
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objContato.sCgc)) > 0 Then
        'Verifica se tem um cliente com o mesmo CGC e dá aviso
        lErro = CF("FilialCliente_Testa_CGC", 0, 0, objContato.sCgc)
        If lErro <> SUCESSO Then gError 180632
    End If
  
    lErro = Trata_Alteracao(objContato, objContato.lCodigo)
    If lErro <> SUCESSO Then gError 180633
    
    'Grava o Contato no BD
    lErro = CF("Contato_Grava", objContato, colEndereco)
    If lErro <> SUCESSO Then gError 180634

'    'gravar anotacao, se houver
'    If Not (gobjAnotacao Is Nothing) Then
'
'        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
'
'            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_CLIENTE
'            gobjAnotacao.sID = CStr(objContato.lCodigo)
'            gobjAnotacao.dtDataAlteracao = gdtDataHoje
'
'            lErro = CF("Anotacoes_Grava", gobjAnotacao)
'            If lErro <> SUCESSO Then gError 180635
'
'        End If
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

        Case 180622
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONTATO_NAO_PREENCHIDO", gErr)

        Case 180623
            Call Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 180624
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 180625
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CONTATO_NAO_PREENCHIDO", gErr)

        Case 180626, 180627, 180628
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case 180629
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_VALOR_CORRESPONDENTE", gErr, GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))

        Case 180630, 180631, 180632, 180633, 180634, 180635 'Tratados nas Rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180636)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 180637

    'Limpa a Tela
    lErro = Limpa_Tela_Contatos()
    If lErro <> SUCESSO Then gError 180638
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 180637, 180638

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180639)

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
            If lErro <> SUCESSO Then gError 180640
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 180641
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            gError 180642

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case gErr

        Case 180640, 180641

        Case 180642
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180643)

    End Select

    Exit Sub

End Sub

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
    If lErro <> SUCESSO Then gError 180644

    'Formata
    ComissaoVendas.Text = Format(ComissaoVendas.Text, "Fixed")

    Exit Sub

Erro_ComissaoVendas_Validate:

    Cancel = True

    Select Case gErr

        Case 180644
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180645)
    
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
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 180646

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 180647

        'Passa a conta para o campo na tela
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 180648

        'conta não cadastrada
        If lErro = 5700 Then gError 180649

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 180648, 180646

        Case 180649
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)

        Case 180647
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180650)

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
    If lErro <> SUCESSO Then gError 180651

    Exit Sub

Erro_DataUltVisita_Validate:

    Cancel = True

    Select Case gErr

        Case 180651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180652)

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
    If lErro <> SUCESSO Then gError 180653

    Desconto.Text = Format(sDesconto, "Fixed")

    Exit Sub

Erro_Desconto_Validate:

    Cancel = True

    Select Case gErr

        Case 180653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180654)

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
    
    'Inicializa variávies AdmEvento
    Set objEventoTipoContato = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoCondicaoPagto = New AdmEvento
    Set objEventoCobrador = New AdmEvento
    Set objEventoMensagem = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoTranspRedespacho = New AdmEvento
    Set objEventoPadraoCobranca = New AdmEvento
    Set objEventoContato = New AdmEvento
    
    RazaoSocial.MaxLength = STRING_CLIENTE_RAZAO_SOCIAL

    'Inicializa variáveis de frame
    iFrame2Atual = 0
    iFrameAtual = 1
    
    'Configura Endereço Principal como visível
    OpcaoEndereco(0).Value = True
    OpcaoEndereco(1).Value = False
    OpcaoEndereco(2).Value = False

    Set objTela = Me
    
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then gError 180657

    'Prepara as Combos e Grid do Tab de Identificacao
    lErro = Inicializa_Tab_Identificacao()
    If lErro <> SUCESSO Then gError 180657
    
    'Prepara as Combos do Tab de Dados Financeiros
    lErro = Inicializa_Tab_DadosFinanceiros()
    If lErro <> SUCESSO Then gError 180658
    
    'Prepara as Combos e a Mascara da Conta Contabil do Tab de Vendas
    lErro = Inicializa_Tab_Vendas()
    If lErro <> SUCESSO Then gError 180660

    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then gError 180661
    
'    If giTipoVersao = VERSAO_LIGHT Then
'        TabelaPreco.ListIndex = 0
'    End If
        
    Set objUser = objUserControl
        
    'Prepara as Combos e a Mascara da Conta Contabil do Tab de Vendas
    lErro = CF("Inicializa_Campos_Obrigatorios", objUser)
    If lErro <> SUCESSO Then gError 180662
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 180657 To 180662
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180663)

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

    'Lê cada código e descrição da tabela TiposDeContato
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 180664

    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Tipo.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Tipo.ItemData(Tipo.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Set objGridCategoria = New AdmGrid

    'Inicializa o Grid de Categoria
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 180665

    'Carrega a combobox de Categoria Contato
    lErro = Carrega_ComboCategoriaContato()
    If lErro <> SUCESSO Then gError 180666

    Inicializa_Tab_Identificacao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Identificacao:

    Inicializa_Tab_Identificacao = gErr

    Select Case gErr
        
        Case 180664, 180665, 180666
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180667)
        
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
    If lErro <> SUCESSO Then gError 180668

    'Preenche a ComboBox TabelaPreco com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        TabelaPreco.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'Preenche Combo CondicaoPagto
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao reduzida da tabela CondicoesPagto
    lErro = CF("CondicoesPagto_Le_Recebimento", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 180669

    'Preenche a ComboBox CondicaoPagto com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'Preenche Combo Mensagem
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela Mensagem
    lErro = CF("Cod_Nomes_Le", "Mensagens", "Codigo", "Descricao", STRING_NFISCAL_MENSAGEM, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 180670

    'Preenche a ComboBox Mensagem com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Mensagem.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Mensagem.ItemData(Mensagem.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    'Essa função só faz sentido para versões customizadas. Na versão 2 é uma função vazia
    lErro = CF("Inicializa_Tab_DadosFinanceiros_Usu", objUserControl)
    If lErro <> SUCESSO Then gError 180671
   
    Inicializa_Tab_DadosFinanceiros = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_DadosFinanceiros:
    
    Inicializa_Tab_DadosFinanceiros = gErr
    
    Select Case gErr
        
        Case 180668 To 180671
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180672)
        
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
    If lErro <> SUCESSO Then gError 180675

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
        If lErro <> SUCESSO Then gError 180676
    
        ContaContabil.Mask = sMascaraConta

    Else
       
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 180676
    
        ContaContabil.Mask = sMascaraConta
        
        'Desabilita Conta Contabil
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
    End If
    

    'Preenche Combo Cobrador
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 180677

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
    If lErro <> SUCESSO Then gError 180678

    'Preenche as ComboBox Transportadora e TranspRedespacho com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Transportadora.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoDescricao.iCodigo
        TranspRedespacho.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TranspRedespacho.ItemData(TranspRedespacho.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    'Lê todos os Padroes de Cobranca da tabela PadroesCobranca
    lErro = CF("PadroesCobranca_Le_Todos", colPadroesCobranca)
    If lErro <> SUCESSO Then gError 180679
    
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
    
    Inicializa_Tab_Vendas = gErr
    
    Select Case gErr
    
        Case 180675 To 180679
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180680)
        
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
    If lErro <> SUCESSO Then gError 180684

    'Formata
    LimiteCredito.Text = Format(sLimiteCredito, "Fixed")

    Exit Sub

Erro_LimiteCredito_Validate:

    Cancel = True

    Select Case gErr

        Case 180684

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180685)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180686

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objMensagem.iCodigo = iCodigo

        'Tenta ler Mensagem com esse código no BD
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 180687

        If lErro <> SUCESSO Then gError 180688 'Não encontrou Mensagem no BD

        'Encontrou Mensagem no BD, coloca no Text da Combo
        Mensagem.Text = CStr(objMensagem.iCodigo) & SEPARADOR & objMensagem.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180689

    Exit Sub

Erro_Mensagem_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 180686, 180687

    Case 180688  'Não encontrou Mensagem no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MENSAGEM")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Mensagens", objMensagem)

        End If

    Case 180689

        Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_ENCONTRADA", gErr, Mensagem.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180690)

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

        If Not IniciaLetra(NomeReduzido.Text) Then gError 180691

    End If

    'Preenche os labels Contato existentes nos frames
    For iIndice = 0 To 3

        ContatoLabel(iIndice).Caption = Trim(NomeReduzido.Text)

    Next

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr

        Case 180691
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", gErr, NomeReduzido.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180692)

    End Select

End Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180694

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objPadraoCobranca.iCodigo = iCodigo

        'Tenta ler Padrao Cobranca com esse código no BD
        lErro = CF("PadraoCobranca_Le", objPadraoCobranca)
        If lErro <> SUCESSO And lErro <> 19298 Then gError 180695

        If lErro = 19298 Then gError 180696 'Não encontrou Padrao Cobranca no BD

        'Encontrou Padrao Cobranca no BD, coloca no Text da Combo
        PadraoCobranca.Text = CStr(objPadraoCobranca.iCodigo) & SEPARADOR & objPadraoCobranca.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180697

    Exit Sub

Erro_PadraoCobranca_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 180694, 180695

    Case 180696  'Não encontrou Padrao Cobranca no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PADRAO_COBRANCA")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("PadroesCobranca", objPadraoCobranca)

        End If

    Case 180697
        Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", gErr, PadraoCobranca.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180698)

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

Public Sub Pais_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180704

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 180705
        
        'Não encontrou Regiao Venda BD
        If lErro <> SUCESSO Then gError 180706
        
        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180707

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 180704, 180705

    Case 180706  'Não encontrou RegiaoVenda no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

        If vbMsgRes = vbYes Then
            'Chama a tela RegiaoVenda
            Call Chama_Tela("RegiaoVenda", objRegiaoVenda)

        End If

    Case 180707
        Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", gErr, Regiao.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180708)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180709

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 180710
        
        If lErro <> SUCESSO Then gError 180711 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180712

    Exit Sub

Erro_TabelaPreco_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 180709, 180710

    Case 180711  'Não encontrou Tabela de Preço no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then

            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            
        End If

    Case 180712
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180713)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Click()

Dim lErro As Long
Dim objTipoContato As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Tipo_Click

    'Configura flag de Alteração
    iAlterado = REGISTRO_ALTERADO

    'Se não foi selecionado Tipo, sai
    If Tipo.ListIndex = -1 Then Exit Sub

    objTipoContato.iCodigo = Tipo.ItemData(Tipo.ListIndex)

    'Lê TipoContato no BD
    lErro = CF("TipoDeCliente_Le", objTipoContato)
    If lErro <> SUCESSO And lErro <> 28943 Then gError 180714

    'Tipo não está cadastrado
    If lErro = 28943 Then gError 180715

    lErro = Exibe_Dados_TipoContato(objTipoContato)
    If lErro <> SUCESSO Then gError 180716

    Exit Sub

Erro_Tipo_Click:

    Select Case gErr

        Case 180714, 180716
            Tipo.SetFocus

        Case 180715  'Tipo não está cadastrado

            'Envia aviso perguntando se deseja cadastrar novo Tipo de Contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoCliente", objTipoContato)

            Else
                Tipo.SetFocus

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180717)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoContato As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Tipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180718

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoContato.iCodigo = iCodigo

        'Tenta ler TipoContato com esse código no BD
        lErro = CF("TipoDeCliente_Le", objTipoContato)
        If lErro <> SUCESSO And lErro <> 28943 Then gError 180719

        'Não encontrou Tipo Contato no BD
        If lErro = 28943 Then gError 180720

        'Exibe dados de TipoContato na tela
        Tipo.Text = CStr(iCodigo) & SEPARADOR & objTipoContato.sDescricao

        lErro = Exibe_Dados_TipoContato(objTipoContato)
        If lErro <> SUCESSO Then gError 180721

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180722

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 180718, 180719, 180721 'Já tratado na rotina chamada

        Case 180720 'Não encontrou Tipo Contato no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a tela de TiposDeContatos
                Call Chama_Tela("TipoCliente", objTipoContato)

            End If

        Case 180722
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOContato_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180723)

    End Select

    Exit Sub

End Sub

Public Sub TipoContatoLabel_Click()

'"""""""""""""""""""""""""""""""""""""""""""
'"      ROTINAS DAS TELAS DE BROWSE        "
'"""""""""""""""""""""""""""""""""""""""""""

'BROWSE TIPO_Contato :

Dim objTipoContato As New ClassTipoCliente
Dim colSelecao As New Collection
    
    'Se o tipo estiver preenchido então passa para objTipoContato
    If Len(Trim(Tipo.Text)) > 0 Then objTipoContato.iCodigo = Codigo_Extrai(Tipo.Text)
    
    'Chama a tela que lista os Contatos
    Call Chama_Tela("TipoClienteLista", colSelecao, objTipoContato, objEventoTipoContato)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180724

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180725
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 180726
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180727

    Exit Sub

Erro_Transportadora_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180724, 180725

        Case 180726  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 180727
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180728)

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
            If lErro <> SUCESSO Then gError 180729
            
            lErro = CF("Vendedor_Comissao_Customiza", objVendedor, dPercComissao)
            If lErro <> SUCESSO Then gError 180730
            
            If Len(Trim(ComissaoVendas.Text)) = 0 And dPercComissao <> 0 Then ComissaoVendas.Text = CStr(100 * dPercComissao)

            If objVendedor.iAtivo = DESMARCADO Then gError 180731
            
        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 180729, 180730   'Tratado na rotina chamada
        
        Case 180731
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180732)
    
    End Select

End Sub

Function Limpa_Tela_Contatos() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Contatos

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
    
    Ativo.Value = MARCADO
    
    Bloqueado.Value = DESMARCADO 'Inserido por Wagner

    Regiao.Text = ""
    Cobrador.Text = ""
    TipoFrete.ListIndex = -1
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False

    'Limpa os Labels Contato
    For iIndice = 0 To 3
        ContatoLabel(iIndice).Caption = ""
    Next

    'Escolhe a opção Endereco Principal
    OpcaoEndereco(0).Value = True
    OpcaoEndereco(1).Value = False
    OpcaoEndereco(2).Value = False

    'Torna visível a Frame do Endereço Principal
    Frame2(iFrame2Atual).Visible = False
    Frame2(0).Visible = True
    iFrame2Atual = 0

    Call Grid_Limpa(objGridCategoria)
   
    'Limpa o código de identificação
    Codigo.Text = ""
    
    Call gobjTabEnd.Limpa_Tela
    
    'Set gobjAnotacao = Nothing

    Limpa_Tela_Contatos = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_Contatos:

    Limpa_Tela_Contatos = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180733)

    End Select
    
    Exit Function
        
End Function

Function Trata_Parametros(Optional objContato As ClassContatos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver Contato passado como parâmetro, exibe seus dados
    If Not (objContato Is Nothing) Then

        'Se Codigo é positivo
        If objContato.lCodigo > 0 Then

            'Lê Contato no BD a partir do código
            lErro = CF("Contato_Le", objContato)
            If lErro <> SUCESSO And lErro <> 180548 Then gError 180734

            'Se não encontrou o Contato no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o código
                lErro = Limpa_Tela_Contatos()
                If lErro <> SUCESSO Then gError 180735
                
                Codigo.Text = CStr(objContato.lCodigo)

            Else  'Encontrou Contato no BD

                'Exibe os dados do Contato
                lErro = Exibe_Dados_Contato(objContato)
                If lErro <> SUCESSO Then gError 180736

            End If

        'se Nome Reduzido está preenchido
        ElseIf Len(Trim(objContato.sNomeReduzido)) > 0 Then

            'Lê Contato no BD a partir do Nome Reduzido
            lErro = CF("Contato_Le_NomeReduzido", objContato)
            If lErro <> SUCESSO And lErro <> 180745 Then gError 180737

            'Se não encontrou o Contato no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o NomeReduzido
                lErro = Limpa_Tela_Contatos()
                If lErro <> SUCESSO Then gError 180738
                
                NomeReduzido.Text = CStr(objContato.sNomeReduzido)

            Else  'Encontrou Contato no BD

                'Exibe os dados do Contato
                lErro = Exibe_Dados_Contato(objContato)
                If lErro <> SUCESSO Then gError 180739

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 180734 To 180739  'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180740)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Exibe_Dados_Contato(objContato As ClassContatos) As Long
'Exibe os dados de Contato na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_Contato

    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180551 Then gError 180750

    'Verifica se Contato não está cadastrado
    If lErro = 180551 Then gError 180751

    'Preenche os campos Contato existentes nos frames
    For iIndice = 0 To 3
        ContatoLabel(iIndice).Caption = objContato.sNomeReduzido
    Next

    'TAB IDENTIFICACAO :
    lErro = Exibe_Dados_Contato_Identificacao(objContato)
    If lErro <> SUCESSO Then gError 180752
    
    'TAB DADOS FINANCEIROS :
    Call Exibe_Dados_Contato_DadosFinanceiros(objContato)
    
    'TAB INSCRICOES :
    Call Exibe_Dados_Contato_Inscricoes(objContato)
    
    'TAB ENDERECOS :
    lErro = Exibe_Dados_Contato_Enderecos(objContato)
    If lErro <> SUCESSO Then gError 180753
    
    'TAB VENDAS:
    lErro = Exibe_Dados_Contato_Vendas(objContato)
    If lErro <> SUCESSO Then gError 180754
        
    iAlterado = 0

    Exibe_Dados_Contato = SUCESSO

    Exit Function

Erro_Exibe_Dados_Contato:
    
    Exibe_Dados_Contato = gErr

    Select Case gErr
        
        Case 180750
        
        Case 180751
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)
        
        Case 180752, 180753, 180754
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180755)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Contato_Identificacao(objContato As ClassContatos) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoContato As New ClassTipoCliente
Dim objFilialContato As New ClassFilialContato
Dim colCategoriaItem As New Collection
Dim iCodigo As Integer

On Error GoTo Erro_Exibe_Dados_Contato_Identificacao
    
    Ativo.Value = objContato.iAtivo
    Codigo.Text = CStr(objContato.lCodigo)
    RazaoSocial.Text = objContato.sRazaoSocial
    NomeReduzido.Text = objContato.sNomeReduzido

    If objContato.iTipo = 0 Then
        Tipo.Text = ""
    Else
        'Coloca Tipo no Text
        Tipo.Text = CStr(objContato.iTipo)

        'Tenta selecionar
        lErro = Combo_Seleciona(Tipo, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And 6731 Then gError 180756

        If lErro <> SUCESSO Then

            objTipoContato.iCodigo = objContato.iTipo

            'Tenta ler TipoContato com esse código no BD
            lErro = CF("TipoDeCliente_Le", objTipoContato)
            If lErro <> SUCESSO And lErro <> 28943 Then gError 180757

            If lErro = 28943 Then gError 180758 'Não encontrou Tipo Contato no BD

            'Insere na Combo
            Tipo.AddItem CStr(objTipoContato.iCodigo) & SEPARADOR & objTipoContato.sDescricao
            Tipo.ItemData(Tipo.NewIndex) = objTipoContato.iCodigo

            'Seleciona na Combo
            lErro = Combo_Seleciona(Tipo, iCodigo)
            If lErro <> SUCESSO Then gError 180759

        End If

    End If

    'CATEGORIAS

    'Lê as categoria deste Contato
    objFilialContato.lCodContato = objContato.lCodigo
    objFilialContato.iCodFilial = FILIAL_MATRIZ
    
    lErro = CF("FilialContato_Le_Categorias", objFilialContato, colCategoriaItem)
    If lErro <> SUCESSO Then gError 180760

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

    Observacao.Text = objContato.sObservacao

    Exibe_Dados_Contato_Identificacao = gErr
    
    Exit Function
    
Erro_Exibe_Dados_Contato_Identificacao:

    Exibe_Dados_Contato_Identificacao = SUCESSO
    
    Select Case gErr
        
        Case 180756, 180757, 180760

        Case 180758
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_Contato_NAO_CADASTRADO", gErr, objTipoContato.iCodigo)

        Case 180759
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_SELECIONADO", gErr, objTipoContato.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180761)

    End Select

    Exit Function

End Function

Private Sub Exibe_Dados_Contato_DadosFinanceiros(objContato As ClassContatos)
'Exibe os dados financeiros do Contato

Dim bCancel As Boolean
    
    LimiteCredito.Text = CStr(objContato.dLimiteCredito)

    If objContato.iBloqueado = MARCADO Then
        Bloqueado.Value = vbChecked
    Else
        Bloqueado.Value = vbUnchecked
    End If

    If objContato.iTabelaPreco = 0 Then
        TabelaPreco.Text = ""
    Else
        TabelaPreco.Text = CStr(objContato.iTabelaPreco)
        Call TabelaPreco_Validate(bCancel)
    End If

    If objContato.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objContato.iCondicaoPagto)
        Call CondicaoPagto_Validate(bCancel)
    End If

    If objContato.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objContato.iCodMensagem)
        Call Mensagem_Validate(bCancel)
    End If

    Desconto.Text = CStr(100 * objContato.dDesconto)

End Sub

Private Sub Exibe_Dados_Contato_Inscricoes(objContato As ClassContatos)
'Exibe as inscrições do Contato
    
    RG.Text = objContato.sRG
    CGC.Text = objContato.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objContato.sInscricaoEstadual
    InscricaoMunicipal.Text = objContato.sInscricaoMunicipal
    Observacao2.Text = objContato.sObservacao2
    
    InscricaoSuframa.PromptInclude = False
    InscricaoSuframa.Text = objContato.sInscricaoSuframa
    InscricaoSuframa.PromptInclude = True
    
End Sub

Private Function Exibe_Dados_Contato_Enderecos(objContato As ClassContatos) As Long
'Exibe os Endereços do Contato

Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco
Dim objCliente As New ClassCliente

On Error GoTo Erro_Exibe_Dados_Contato_Enderecos

    objCliente.lEndereco = objContato.lEndereco
    objCliente.lEnderecoCobranca = objContato.lEnderecoCobranca
    objCliente.lEnderecoEntrega = objContato.lEnderecoEntrega

    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
    If lErro <> SUCESSO Then gError 180767

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 180767

    Exibe_Dados_Contato_Enderecos = SUCESSO
    
    Exit Function
    
Erro_Exibe_Dados_Contato_Enderecos:
    
    Exibe_Dados_Contato_Enderecos = gErr
    
    Select Case gErr
            
        Case 180767
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180768)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Contato_Vendas(objContato As ClassContatos) As Long
'Exibe dados de Vendas do Contato

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_Contato_Vendas

    If objContato.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objContato.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If
    
    'Tipo de Frete
    If objContato.iTipoFrete <> -1 Then
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objContato.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoFrete.ListIndex = -1
    End If
    
    ContaContabil.PromptInclude = False
    If Len(Trim(objContato.sContaContabil)) = 0 Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objContato.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 180769

        ContaContabil.Text = sContaEnxuta
    End If
    ContaContabil.PromptInclude = True

    If objContato.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objContato.dComissaoVendas)
    End If

    If objContato.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objContato.iRegiao)
        Call Regiao_Validate(bCancel)
    End If

    If objContato.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objContato.iCodCobrador)
        Call Cobrador_Validate(bCancel)
    End If

    If objContato.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objContato.iFreqVisitas)
    End If

    If (objContato.dtDataUltVisita) = DATA_NULA Then
        DataUltVisita.PromptInclude = False
        DataUltVisita.Text = ""
        DataUltVisita.PromptInclude = True
    Else
        DataUltVisita.Text = Format(objContato.dtDataUltVisita, "dd/mm/yy")
    End If

    If objContato.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objContato.iCodTransportadora)
        Call Transportadora_Validate(bCancel)
    End If

    If objContato.iCodPadraoCobranca = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objContato.iCodPadraoCobranca)
        Call PadraoCobranca_Validate(bCancel)
    End If

    Guia.Text = objContato.sGuia
    
    If objContato.iCodTranspRedesp = 0 Then
        TranspRedespacho.Text = ""
    Else
        TranspRedespacho.Text = CStr(objContato.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If

    RedespachoCli.Value = objContato.iDetPagFrete

    Exibe_Dados_Contato_Vendas = SUCESSO
    
    Exit Function
    
Erro_Exibe_Dados_Contato_Vendas:

    Exibe_Dados_Contato_Vendas = gErr
    
    Select Case gErr
            
        Case 180769
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objContato.sContaContabil)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180770)

    End Select

    Exit Function

End Function

Function Exibe_Dados_TipoContato(objTipoContato As ClassTipoCliente) As Long
'Exibe dados do Tipo de Contato na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim colCategoriaItem As New Collection
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_TipoContato

    LimiteCredito.Text = CStr(objTipoContato.dLimiteCredito)

    If objTipoContato.iTabelaPreco = 0 Then
        TabelaPreco.Text = ""
    Else
        TabelaPreco.Text = CStr(objTipoContato.iTabelaPreco)
        Call TabelaPreco_Validate(bCancel)
    End If

    If objTipoContato.iCondicaoPagto = 0 Then
        CondicaoPagto.Text = ""
    Else
        CondicaoPagto.Text = CStr(objTipoContato.iCondicaoPagto)
        Call CondicaoPagto_Validate(bCancel)
    End If

    If objTipoContato.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objTipoContato.iCodMensagem)
        Call Mensagem_Validate(bCancel)
    End If

    If objTipoContato.dDesconto = 0# Then
        Desconto.Text = ""
    Else
        Desconto.Text = CStr(100 * objTipoContato.dDesconto)
    End If

    ContaContabil.PromptInclude = False

    If objTipoContato.sContaContabil = "" Then
        ContaContabil.Text = ""
    Else
        lErro = Mascara_RetornaContaEnxuta(objTipoContato.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 180771

        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    If objTipoContato.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objTipoContato.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If

    If objTipoContato.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objTipoContato.dComissaoVendas)
    End If

    If objTipoContato.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objTipoContato.iFreqVisitas)
    End If

    If objTipoContato.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objTipoContato.iCodTransportadora)
        Call Transportadora_Validate(bCancel)
    End If

    If objTipoContato.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objTipoContato.iCodCobrador)
        Call Cobrador_Validate(bCancel)
    End If

    If objTipoContato.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objTipoContato.iRegiao)
        Call Regiao_Validate(bCancel)
    End If
    
    If objTipoContato.iPadraoCobranca = 0 Then
        PadraoCobranca.Text = ""
    Else
        PadraoCobranca.Text = CStr(objTipoContato.iPadraoCobranca)
        Call PadraoCobranca_Validate(bCancel)
    End If
    
    'Lê as Categorias de Tipo de Contato no BD
    lErro = CF("TipoDeClienteCategorias_Le", objTipoContato, colCategoriaItem)
    If lErro <> SUCESSO Then gError 180772

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

    Exibe_Dados_TipoContato = SUCESSO

    Exit Function

Erro_Exibe_Dados_TipoContato:

    Exibe_Dados_TipoContato = gErr

    Select Case gErr

        Case 180771
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objTipoContato.sContaContabil)

        Case 180772

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180773)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Contato(objContato As ClassContatos) As Long
'Lê os dados que estão na tela de Contatos e coloca em objContato

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objVendedor As New ClassVendedor
Dim objFilialConCategoria As New ClassFilialConCategoria
Dim objTabelaPreco As New ClassTabelaPreco
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objMensagem As New ClassMensagem
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objCobrador As New ClassCobrador
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Le_Dados_Contato

    'IDENTIFICACAO :

    If Len(Trim(Codigo.Text)) > 0 Then objContato.lCodigo = CLng(Codigo.Text)

    objContato.sRazaoSocial = Trim(RazaoSocial.Text)
    objContato.sNomeReduzido = Trim(NomeReduzido.Text)

    If Len(Trim(Tipo.Text)) > 0 Then objContato.iTipo = Codigo_Extrai(Tipo.Text)

    objContato.sObservacao = Trim(Observacao.Text)
    
    objContato.iAtivo = Ativo.Value
    
    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialConCategoria = New ClassFilialConCategoria

            objFilialConCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialConCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            objContato.colCategoriaItem.Add objFilialConCategoria

        End If

    Next

    'DADOS FINANCEIROS :

    If Bloqueado.Value = vbChecked Then
        objContato.iBloqueado = MARCADO
    Else
        objContato.iBloqueado = DESMARCADO
    End If

    If Len(Trim(LimiteCredito.Text)) > 0 Then objContato.dLimiteCredito = CDbl(LimiteCredito.Text)
    If Len(Trim(Desconto.Text)) > 0 Then objContato.dDesconto = CDbl(Desconto.Text) / 100

    If Len(Trim(TabelaPreco.Text)) > 0 Then objContato.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)

    'Verifica se tabela de preço existe
    If objContato.iTabelaPreco <> 0 Then
        objTabelaPreco.iCodigo = objContato.iTabelaPreco
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 180774

        'Não encontrou a tabela de preço ==> erro
        If lErro = 28004 Then gError 180775

        If objTabelaPreco.iAtivo = DESMARCADO Then gError 213837

    End If

    If Len(Trim(CondicaoPagto.Text)) > 0 Then objContato.iCondicaoPagto = Codigo_Extrai(CondicaoPagto.Text)

    'Verifica se a condição de pagamento existe
    If objContato.iCondicaoPagto <> 0 Then
        objCondicaoPagto.iCodigo = objContato.iCondicaoPagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 180776

        'Não encontrou a condição de pagamento ==> erro
        If lErro = 19205 Then gError 180777

    End If

    If Len(Trim(Mensagem.Text)) > 0 Then objContato.iCodMensagem = Codigo_Extrai(Mensagem.Text)

    'Verifica se mensagem existe
    If objContato.iCodMensagem <> 0 Then
        objMensagem.iCodigo = objContato.iCodMensagem
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 180778

        'Não encontrou a mensagem ==> erro
        If lErro = 19234 Then gError 180779

    End If

    'INSCRICOES :
    
    objContato.sRG = Trim(RG.Text)
    objContato.sCgc = Trim(CGC.Text)
    objContato.sInscricaoEstadual = Trim(InscricaoEstadual.Text)
    objContato.sInscricaoMunicipal = Trim(InscricaoMunicipal.Text)
    objContato.sObservacao2 = Trim(Observacao2.Text)
    objContato.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)
    
    'VENDAS :

    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 180780

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 180781

        objContato.iVendedor = objVendedor.iCodigo

    End If

    sConta = String(STRING_CONTA, 0)

    'Formata Conta contábil
    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 180782

    If iContaPreenchida = CONTA_VAZIA Then
        objContato.sContaContabil = ""
    Else
        objContato.sContaContabil = sConta
    End If

    If Len(Trim(ComissaoVendas.Text)) > 0 Then objContato.dComissaoVendas = CDbl(ComissaoVendas.Text) / 100

    If Len(Trim(Regiao.Text)) > 0 Then objContato.iRegiao = Codigo_Extrai(Regiao.Text)

    If Len(Trim(PadraoCobranca.Text)) > 0 Then objContato.iCodPadraoCobranca = Codigo_Extrai(PadraoCobranca.Text)
    
    'Verifica se região venda existe
    If objContato.iRegiao <> 0 Then
        objRegiaoVenda.iCodigo = objContato.iRegiao
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 180783

        'Não encontrou a região venda ==> erro
        If lErro = 16137 Then gError 180784

    End If

    If Len(Trim(Cobrador.Text)) > 0 Then objContato.iCodCobrador = Codigo_Extrai(Cobrador.Text)

    'Verifica se cobrador existe
    If objContato.iCodCobrador <> 0 Then
        objCobrador.iCodigo = objContato.iCodCobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 180785

        'Não encontrou o cobrador ==> erro
        If lErro = 19294 Then gError 180786

    End If

    If Len(Trim(FreqVisitas.Text)) > 0 Then objContato.iFreqVisitas = CInt(FreqVisitas.Text)

    If Trim(DataUltVisita.Text) = "/  /" Then
        objContato.dtDataUltVisita = DATA_NULA
    Else
        objContato.dtDataUltVisita = CDate(DataUltVisita.Text)
    End If

    'Tipo de Frete
    If TipoFrete.ListIndex <> -1 Then
        If TipoFrete.ItemData(TipoFrete.ListIndex) = TIPO_CIF Then
            objContato.iTipoFrete = TIPO_CIF
        Else
            objContato.iTipoFrete = TIPO_FOB
        End If
    Else
        objContato.iTipoFrete = -1
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then objContato.iCodTransportadora = Codigo_Extrai(Transportadora.Text)

    'Verifica se transportadora existe
    If objContato.iCodTransportadora <> 0 Then
        objTransportadora.iCodigo = objContato.iCodTransportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180787

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 180788

    End If

    objContato.sGuia = Trim(Guia.Text)
    
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objContato.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)

    'Verifica se transportadora existe
    If objContato.iCodTranspRedesp <> 0 Then
        objTransportadora.iCodigo = objContato.iCodTranspRedesp
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180789

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 180790

    End If

    objContato.iDetPagFrete = RedespachoCli.Value

    Le_Dados_Contato = SUCESSO

    Exit Function

Erro_Le_Dados_Contato:

    Le_Dados_Contato = gErr

    Select Case gErr
        
        'Erro tratado na rotina chamada
        Case 180774, 180776, 180778, 180780, 180782, 180783, 180785, 180787, 180789

        Case 180775
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", gErr, objTabelaPreco.iCodigo)

        Case 180777
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 180779
            Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", gErr, objMensagem.iCodigo)

        Case 180781
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 180784
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case 180786
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)

        Case 180788, 180790
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)

        Case ERRO_SEM_MENSAGEM
            
        Case 213837
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_INATIVA", gErr, objTabelaPreco.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180791)

    End Select

    Exit Function

End Function

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

'RETORNO TIPO_Contato :
Private Sub objEventoTipoContato_evSelecao(obj1 As Object)

Dim objTipoContato As ClassTipoCliente
Dim lErro As Long
Dim bCancel As Boolean

    Set objTipoContato = obj1

    'Preenche Text da Combo Tipo
    Tipo.Text = CStr(objTipoContato.iCodigo)
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

    Vendedor.SetFocus
    
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
        If lErro <> SUCESSO Then gError 180794

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 180794
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180795)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objContato As New ClassContatos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Contatos"

    'Lê os dados da Tela Contatos
    lErro = Le_Dados_Contato(objContato)
    If lErro <> SUCESSO Then gError 180796

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objContato.lCodigo, 0, "Codigo"
    colCampoValor.Add "RazaoSocial", objContato.sRazaoSocial, STRING_CLIENTE_RAZAO_SOCIAL, "RazaoSocial"
    colCampoValor.Add "NomeReduzido", objContato.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Tipo", objContato.iTipo, 0, "Tipo"
    colCampoValor.Add "Observacao", objContato.sObservacao, STRING_CLIENTE_OBSERVACAO, "Observacao"
    colCampoValor.Add "LimiteCredito", objContato.dLimiteCredito, 0, "LimiteCredito"
    colCampoValor.Add "CondicaoPagto", objContato.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Desconto", objContato.dDesconto, 0, "Desconto"
    colCampoValor.Add "CodMensagem", objContato.iCodMensagem, 0, "CodMensagem"
    colCampoValor.Add "TabelaPreco", objContato.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "CodPadraoCobranca", objContato.iCodPadraoCobranca, 0, "CodPadraoCobranca"
    colCampoValor.Add "Ativo", objContato.iAtivo, 0, "Ativo"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 180796

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180797)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objContato As New ClassContatos
Dim objFilialContato As New ClassFilialContato

On Error GoTo Erro_Tela_Preenche

    objContato.lCodigo = colCampoValor.Item("Codigo").vValor

    If objContato.lCodigo <> 0 Then

        'Carrega objContato com os dados passados em colCampoValor
        objContato.sRazaoSocial = colCampoValor.Item("RazaoSocial").vValor
        objContato.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objContato.iTipo = colCampoValor.Item("Tipo").vValor
        objContato.sObservacao = colCampoValor.Item("Observacao").vValor
        objContato.dLimiteCredito = colCampoValor.Item("LimiteCredito").vValor
        objContato.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objContato.dDesconto = colCampoValor.Item("Desconto").vValor
        objContato.iCodMensagem = colCampoValor.Item("CodMensagem").vValor
        objContato.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
        objContato.iCodPadraoCobranca = colCampoValor.Item("CodPadraoCobranca").vValor
        objContato.iAtivo = colCampoValor.Item("Ativo").vValor
        
        'Inicializa objFilialContato
        objFilialContato.lCodContato = objContato.lCodigo
        objFilialContato.iCodFilial = FILIAL_MATRIZ

        'Lê o restante dos dados do Contato na tabela de Filiais
        lErro = CF("FilialContato_Le", objFilialContato)
        If lErro <> SUCESSO And lErro <> 180801 Then gError 180803

        'Carrega os restantes dos dados do Contato contidos em objFilialContato
        objContato.sRG = objFilialContato.sRG
        objContato.sCgc = objFilialContato.sCgc
        objContato.sInscricaoEstadual = objFilialContato.sInscricaoEstadual
        objContato.sInscricaoMunicipal = objFilialContato.sInscricaoMunicipal
        objContato.sInscricaoSuframa = objFilialContato.sInscricaoSuframa
        objContato.lEndereco = objFilialContato.lEndereco
        objContato.lEnderecoEntrega = objFilialContato.lEnderecoEntrega
        objContato.lEnderecoCobranca = objFilialContato.lEnderecoCobranca
        objContato.iTipoFrete = objFilialContato.iTipoFrete
        objContato.iCodTransportadora = objFilialContato.iCodTransportadora
        objContato.iCodTranspRedesp = objFilialContato.iCodTranspRedesp
        objContato.iDetPagFrete = objFilialContato.iDetPagFrete
        objContato.sGuia = objFilialContato.sGuia
        objContato.sObservacao2 = objFilialContato.sObservacao
        objContato.sContaContabil = objFilialContato.sContaContabil
        objContato.iVendedor = objFilialContato.iVendedor
        objContato.dComissaoVendas = objFilialContato.dComissaoVendas
        objContato.iRegiao = objFilialContato.iRegiao
        objContato.iFreqVisitas = objFilialContato.iFreqVisitas
        objContato.dtDataUltVisita = objFilialContato.dtDataUltVisita
        objContato.iCodCobrador = objFilialContato.iCodCobrador

        'Exibe o Contato na Tela
        lErro = Exibe_Dados_Contato(objContato)
        If lErro <> SUCESSO Then gError 180804

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 180803, 180804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180805)

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
    
    'Set gobjAnotacao = Nothing
    
    'Descarrega variáveis globais tipo AdmEvento
    Set objEventoTipoContato = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoCondicaoPagto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoCobrador = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoPadraoCobranca = Nothing
    Set objEventoContato = Nothing

    'Descarrega variável global
    Set objGridCategoria = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

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
    objGridInt.colCampo.Add (ComboCategoriaContato.Name)
    objGridInt.colCampo.Add (ComboCategoriaContatoItem.Name)

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

Private Function Trata_ComboCategoriaContatoItem()

Dim lErro As Long
Dim objCategoriaContato As New ClassCategoriaCliente
Dim iIndice As Integer, sValor As String

On Error GoTo Erro_Trata_ComboCategoriaContatoItem

    'Armazena o Item em sValor
    sValor = ComboCategoriaContatoItem.Text

    'Limpa a Combo
    ComboCategoriaContatoItem.Clear

    'Retorna o Item para a Combo
    ComboCategoriaContatoItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(Trim(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col))) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaContato.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaContatoItem(objCategoriaContato)
        If lErro <> SUCESSO Then gError 180806

    End If

    'Seleciona o Item na List da Combo
    For iIndice = 0 To ComboCategoriaContatoItem.ListCount - 1
        If ComboCategoriaContatoItem.List(iIndice) = GridCategoria.Text Then
            ComboCategoriaContatoItem.ListIndex = iIndice
            Exit For
        End If
    Next

    Trata_ComboCategoriaContatoItem = SUCESSO

    Exit Function

Erro_Trata_ComboCategoriaContatoItem:

    Trata_ComboCategoriaContatoItem = gErr

    Select Case gErr

        Case 180806

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180807)

    End Select

End Function

Private Function Carrega_ComboCategoriaContatoItem(objCategoriaContato As ClassCategoriaCliente) As Long
'Carrega a Combo CategoriaContatoItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaContatoItem As ClassCategoriaClienteItem

On Error GoTo Erro_Carrega_ComboCategoriaContatoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaContato, colItensCategoria)
    If lErro <> SUCESSO Then gError 180808

    'Insere na combo CategoriaContatoItem
    For Each objCategoriaContatoItem In colItensCategoria

        'Insere na combo CategoriaContato
        ComboCategoriaContatoItem.AddItem objCategoriaContatoItem.sItem

    Next

    Carrega_ComboCategoriaContatoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaContatoItem:

    Carrega_ComboCategoriaContatoItem = gErr

    Select Case gErr

        Case 180808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180809)

    End Select

    Exit Function

End Function

Public Sub Label1_Click()

Dim objContato As New ClassContatos
Dim colSelecao As Collection

    'Preenche NomeReduzido com o Contato da tela
    If Len(Trim(Codigo.Text)) > 0 Then objContato.lCodigo = CDbl(Codigo.Text)

    'Chama Tela ContatoLista
    Call Chama_Tela("ContatosLista", colSelecao, objContato, objEventoContato)

End Sub

Public Sub Label3_Click()

Dim objContato As New ClassContatos
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_Label3_Click

    objContato.sNomeReduzido = NomeReduzido.Text

    'Chama Tela ContatoLista
    Call Chama_Tela("ContatosLista", colSelecao, objContato, objEventoContato)

    Exit Sub

Erro_Label3_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180810)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContato_evSelecao(obj1 As Object)

Dim objContato As ClassContatos
Dim bCancel As Boolean

    Set objContato = obj1

    'Executa o Validate
    Call Contato_Traz_Tela(objContato.lCodigo)

    Me.Show

    Exit Sub

End Sub

Public Sub Contato_Traz_Tela(ByVal lCodigo As Long)

Dim lErro As Long
Dim objContato As New ClassContatos

On Error GoTo Erro_Contato_Traz_Tela

    'Guarda o valor do código do Contato selecionado na ListBox ContatosList
    objContato.lCodigo = lCodigo

    'Lê o Contato no BD
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180551 Then gError 180811

    'Se Contato não está cadastrado, erro
    If lErro = 180551 Then gError 180812

    'Exibe os dados do Contato
    lErro = Exibe_Dados_Contato(objContato)
    If lErro <> SUCESSO Then gError 180813

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Contato_Traz_Tela:

    Select Case gErr

        Case 180811, 180813

        Case 180812
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180814)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaContato() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaContato As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaContato

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos", colCategorias)
    If lErro <> SUCESSO Then gError 180815

    For Each objCategoriaContato In colCategorias

        'Insere na combo CategoriaContato
        ComboCategoriaContato.AddItem objCategoriaContato.sCategoria

    Next

    Carrega_ComboCategoriaContato = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaContato:

    Carrega_ComboCategoriaContato = gErr

    Select Case gErr

        Case 180815

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180816)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCategoria.Col

            Case GRID_CATEGORIA_COL

                'Chama SaidaCelula de Categoria
                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then gError 180817

            Case GRID_VALOR_COL

                'Chama SaidaCelula de Valor
                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 180818

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 180819

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 180817, 180818

        Case 180819
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180820)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Categoria do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaContato As New ClassCategoriaCliente

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaContato

    'Testa se .Text da combo é igual ao TextMatrix "embaixo"
    iIndice = GridCategoria.Row
    
    'Se não for,
    If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) <> ComboCategoriaContato.Text Then

        If Len(Trim(ComboCategoriaContato.Text)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaContato.sCategoria = ComboCategoriaContato.Text

            'Lê Categoria De Contato no BD
            lErro = CF("CategoriaCliente_Le", objCategoriaContato)
            If lErro <> SUCESSO And lErro <> 28847 Then gError 180821

            'Categoria não está cadastrada
            If lErro = 28847 Then gError 180822

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        'Se já existe a categoria, erro
                        If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaContato.Text Then gError 180824

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
    If lErro <> SUCESSO Then gError 180823

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 180821
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 180822 'Categoria não está cadastrada

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaContato"
                Call Chama_Tela("CategoriaCliente", objCategoriaContato)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 180823
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 180824
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180825)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCategoriaContato As New ClassCategoriaCliente
Dim objCategoriaContatoItem As New ClassCategoriaClienteItem
Dim colItens As New Collection
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaContatoItem

    If Len(Trim(ComboCategoriaContatoItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then gError 180826

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaContatoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaContatoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaContatoItem.sItem = ComboCategoriaContatoItem.Text

            'Lê Item De Categoria De Contato no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaContatoItem)
            If lErro <> SUCESSO And lErro <> 28991 Then gError 180827

            'Item da Categoria não está cadastrado
            If lErro = 28991 Then gError 180828

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 180829

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 180826
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 180827
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 180828 'Item da Categoria não está cadastrado

            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTEITEM")

            If vbMsgRes = vbYes Then

                'Preenche o objeto com a Categoria
                objCategoriaContato.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)

                'Chama a Tela "CategoriaContato"
                Call Chama_Tela("CategoriaCliente", objCategoriaContato, objCategoriaContatoItem.sItem)
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 180829
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180830)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CLIENTES_ID
    Set Form_Load_Ocx = Me
    Caption = "Clientes Futuros"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Contatos"
    
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
            Call TipoContatoLabel_Click
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180831

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180832
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 180833
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180834

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180831, 180832

        Case 180833  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 180834
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180835)

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

Public Sub Bloqueado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoContatos_Click()

Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_BotaoContatos_Click
    
    objClienteContatos.lCliente = StrParaDbl(Codigo.Text)
    objClienteContatos.iFilialCliente = FILIAL_MATRIZ
    
    Call Chama_Tela("ClienteFContatos", objClienteContatos)

    Exit Sub
    
Erro_BotaoContatos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

End Sub
