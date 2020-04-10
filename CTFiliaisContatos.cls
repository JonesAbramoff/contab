VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTFiliaisContatos"
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
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1
Private WithEvents objEventoFilialContato As AdmEvento
Attribute objEventoFilialContato.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Inscricoes = 2
Private Const TAB_Enderecos = 3
Private Const TAB_Vendas = 4

'--- inicio dos properties get dos controles da tela

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

Public Property Get ComboCategoriaContatoItem() As Object
     Set ComboCategoriaContatoItem = objUserControl.Controls("ComboCategoriaContatoItem")
End Property

Public Property Get ComboCategoriaContato() As Object
     Set ComboCategoriaContato = objUserControl.Controls("ComboCategoriaContato")
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

Public Property Get CodContato() As Object
     Set CodContato = objUserControl.Controls("CodContato")
End Property

Public Property Get LabelCodContato() As Object
     Set LabelCodContato = objUserControl.Controls("LabelCodContato")
End Property

Public Property Get LabelCodFilial() As Object
     Set LabelCodFilial = objUserControl.Controls("LabelCodFilial")
End Property

Public Property Get Nome() As Object
     Set Nome = objUserControl.Controls("Nome")
End Property

Public Property Get ContatoLabel() As Object
     Set ContatoLabel = objUserControl.Controls("ContatoLabel")
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
'--- fim dos properties get dos controles da tela

Public Sub LabelCodContato_Click()

Dim colSelecao As Collection
Dim objContato As New ClassContatos

    'Preenche ContatoAte com o Contato da tela
    objContato.lCodigo = StrParaLong(CodContato.Text)

    'Chama Tela ContatosLista
    Call Chama_Tela("ContatosLista", colSelecao, objContato, objEventoContato)

End Sub

Private Sub objEventoContato_evSelecao(obj1 As Object)

Dim objContato As ClassContatos
Dim bCancel As Boolean

    Set objContato = obj1

    'Preenche campo Agente Cobrador
    CodContato.Text = CStr(objContato.lCodigo)

    Call CodContato_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCodFilial_Click()

Dim colSelecao As Collection
Dim objFilialContato As New ClassFilialContato

    'Preenche ContatoAte com o Contato da tela
    objFilialContato.lCodContato = StrParaLong(CodContato.Text)
    objFilialContato.iCodFilial = StrParaInt(CodFilial.Text)

    'Chama Tela ContatosLista
    Call Chama_Tela("FiliaisContatosLista", colSelecao, objFilialContato, objEventoFilialContato)

End Sub

Private Sub objEventoFilialContato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFilialContato As ClassFilialContato

On Error GoTo Erro_objEventoFilialContato_evSelecao

    Set objFilialContato = obj1

    'Tenta ler Filial de Contato com a chave passada em objFilialContato
    lErro = CF("FilialContato_Le", objFilialContato)
    If lErro <> SUCESSO And lErro <> 52591 Then gError 180840

    'Se Filial não existe
    If lErro <> SUCESSO Then gError 180841

    'Exibe dados da Filial na Tela
    lErro = Exibe_Dados_FilialContato(objFilialContato)
    If lErro <> SUCESSO Then gError 180842

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoFilialContato_evSelecao:

    Select Case gErr

        Case 180840, 180842

        Case 180841
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALContato_NAO_CADASTRADA", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180843)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático de Filial
    lErro = CF("FiliaisContatos_Automatico", CLng(CodContato.Text), iCodigo)
    If lErro <> SUCESSO Then gError 180857

    'Coloca código gerado na Tela
    CodFilial.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 180857
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180858)
    
    End Select

    Exit Sub

End Sub

Public Sub Bairro_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilialContato As New ClassFilialContato
Dim objContato As New ClassContatos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(CodContato.Text)) = 0 Then gError 180858

    'Verifica se o código da Filial foi preenchido
    If Len(Trim(CodFilial.Text)) = 0 Then gError 180859

    'Verifica se é Matriz
    If CInt(CodFilial.Text) = FILIAL_MATRIZ Then gError 180860

    objContato.lCodigo = StrParaLong(CodContato.Text)
    objFilialContato.lCodContato = StrParaLong(CodContato.Text)
    objFilialContato.iCodFilial = StrParaInt(CodFilial.Text)

    'Lê os dados do Contato
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 180861

    'Verifica se Contato não está cadastrado
    If lErro <> SUCESSO Then gError 180862

    'Lê os dados da Filial Contato
    lErro = CF("FilialContato_Le", objFilialContato)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 180863

    'Verifica se a Filial Contato não está cadastrada
    If lErro <> SUCESSO Then gError 180864

    'Envia aviso perguntando se realmente deseja excluir Filial Contato
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FILIALCONTATO")

    If vbMsgRes = vbYes Then

        'Exclui Filial de Contato
        lErro = CF("FilialContato_Exclui", objFilialContato)
        If lErro <> SUCESSO Then gError 180865

        'Limpa a Tela
        lErro = Limpa_Tela_FiliaisContatos()
        If lErro <> SUCESSO Then gError 180866
        
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 180858
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONTATO_NAO_PREENCHIDO", gErr)

        Case 180859
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 180860
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCONTATO_EXCLUSAO_MATRIZ", gErr)

        Case 180861, 180863, 180865, 180866

        Case 180862
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)

        Case 180864
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALContato_NAO_CADASTRADA", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180867)

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
    If lErro <> SUCESSO Then gError 180868

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisContatos()
    If lErro <> SUCESSO Then gError 180869
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 180868, 180869

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180870)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar modificações feitas
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 180871

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisContatos()
    If lErro <> SUCESSO Then gError 180872

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 180871, 180872
        
        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180873)

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
        If lErro <> SUCESSO Then gError 180874

        CGC.Format = "000\.000\.000-00; ; ; "
        CGC.Text = CGC.Text

    Case STRING_CGC 'CGC

        'Critica CGC
        lErro = Cgc_Critica(CGC.Text)
        If lErro <> SUCESSO Then gError 180875

        CGC.Format = "00\.000\.000\/0000-00; ; ; "
        CGC.Text = CGC.Text

    Case Else

        gError 180876

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case gErr

        Case 180874, 180875

        Case 180876
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180877)

    End Select


    Exit Sub

End Sub

Public Sub RG_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RG_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180878

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        'Lê Cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 180879

        If lErro <> SUCESSO Then gError 180880 'Não encontrou Cobrador no BD

        'Encontrou Cobrador no BD, coloca no Text da Combo
        Cobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180881

    Exit Sub

Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 180878, 180879

    Case 180880  'Não encontrou Cobrador no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Cobradores", objCobrador)

        End If

    Case 180881

        Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, Cobrador.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180882)

    End Select

    Exit Sub

End Sub

Public Sub CodContato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodContato_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodContato, iAlterado)

End Sub

Public Sub CodContato_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objContato As New ClassContatos
Dim iIndice As Integer

On Error GoTo Erro_CodContato_Validate

    'Verifica se foi preenchido o campo CodContato
    If Len(Trim(CodContato.Text)) = 0 Then Exit Sub

    'Critica se é do tipo Long positivo
    lErro = Long_Critica(CodContato.Text)
    If lErro <> SUCESSO Then gError 180883

    objContato.lCodigo = CLng(CodContato.Text)

    'Lê o Contato
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180548 Then gError 180884

    'Coloca o Nome Reduzido nas labels (mesmo se estiverem vazios)
    For iIndice = 0 To 3
        ContatoLabel(iIndice).Caption = objContato.sNomeReduzido
    Next

    'Verifica se existe
    If lErro <> SUCESSO Then

        'Envia aviso perguntando se deseja cadastrar novo Contato
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO")

        If vbMsgRes = vbYes Then

            'Chama tela de Contatos
            lErro = Chama_Tela("Contatos", objContato)
            If lErro <> SUCESSO Then gError 180885

        Else

            Cancel = True

        End If

    End If

    Exit Sub

Erro_CodContato_Validate:

    Cancel = True

    Select Case gErr

        Case 180883, 180884, 180885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180886)

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
    If lErro <> SUCESSO Then gError 180887

    Exit Sub

Erro_CodFilial_Validate:

    Cancel = True

    Select Case gErr

        Case 180887

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180888)

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
    lErro = Trata_ComboCategoriaContatoItem()
    If lErro <> SUCESSO Then gError 180889

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaContatoItem_GotFocus:

    Select Case gErr

        Case 180889

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180890)

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
        If lErro <> SUCESSO Then gError 180891

    End If

    'Seleciona Item que está no Grid
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

        Case 180891

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180892)

    End Select

End Function

Private Function Carrega_ComboCategoriaContatoItem(objCategoriaContato As ClassCategoriaCliente) As Long
'Carrega os Itens da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaContatoItem As ClassCategoriaClienteItem

On Error GoTo Erro_Carrega_ComboCategoriaContatoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaContato, colItensCategoria)
    If lErro <> SUCESSO Then gError 180893

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

        Case 180893

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180894)

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
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 180895

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 180896

        'Carrega a conta formatada na tela
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True

    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica a conta
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 180897

        If lErro = 5700 Then gError 180898

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 180895, 180897
    
        Case 180896
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
    
        Case 180898
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180899)
    
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
    If lErro <> SUCESSO Then gError 180900

    ComissaoVendas.Text = Format(sComissaoVendas, "Fixed")

    Exit Sub

Erro_ComissaoVendas_Validate:

    Cancel = True

    Select Case gErr

        Case 180900
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180901)
    
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
    If lErro <> SUCESSO Then gError 180903

    Exit Sub

Erro_DataUltVisita_Validate:

    Cancel = True

    Select Case gErr

        Case 180903
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180904)
    
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180912

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 180913

        If lErro <> SUCESSO Then gError 180914 'Não encontrou Regiao Venda BD

        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180915

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180912, 180913

        Case 180914  'Não encontrou RegiaoVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

            If vbMsgRes = vbYes Then
                'Chama a tela de RegiaoVendasDeContatos
                Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
    
            End If
    
        Case 180915
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", gErr, Regiao.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180916)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 180917

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180918
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 180919

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 180920

    Exit Sub

Erro_Transportadora_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 180917, 180918
    
        Case 180919  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")
    
            If vbMsgRes = vbYes Then
    
                Call Chama_Tela("Transportadora", objTransportadora)
    
            End If
    
        Case 180920
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180921)

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
            If lErro <> SUCESSO Then gError 180922

            lErro = CF("Vendedor_Comissao_Customiza", objVendedor, dPercComissao)
            If lErro <> SUCESSO Then gError 180923
            
            If Len(Trim(ComissaoVendas.Text)) = 0 And dPercComissao <> 0 Then ComissaoVendas.Text = CStr(100 * dPercComissao)

            If objVendedor.iAtivo = DESMARCADO Then gError 180924

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 180922, 180923   'Tratado na rotina chamada
                        
        Case 180924
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180925)

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
Dim sCodContato As String
Dim objFilialContato As New ClassFilialContato

On Error GoTo Erro_Filiais_NodeClick

    'Se não deu click em Filial sai da rotina
    If Node.Parent Is Nothing Then Exit Sub

    'Determina chave CodContato-CodFilial associada ao Nó
    sCodFilial = right(Node.Key, Len(Node.Key) - Len(Node.Parent.Key) - 1)
    sCodContato = right(Node.Parent.Key, Len(Node.Parent.Key) - 1)

    objFilialContato.lCodContato = StrParaLong(sCodContato)
    objFilialContato.iCodFilial = StrParaInt(sCodFilial)

    'Tenta ler Filial de Contato com a chave passada em objFilialContato
    lErro = CF("FilialContato_Le", objFilialContato)
    If lErro <> SUCESSO And lErro <> 52591 Then gError 180926

    'Se Filial não existe
    If lErro <> SUCESSO Then gError 180927

    'Exibe dados da Filial na Tela
    lErro = Exibe_Dados_FilialContato(objFilialContato)
    If lErro <> SUCESSO Then gError 180928

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Filiais_NodeClick:

    Select Case gErr

        Case 180926, 180928

        Case 180927
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCONTATO_NAO_CADASTRADA", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180929)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_FiliaisContatos() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sCodigoContato As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FiliaisContatos

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Guarda o código de Contato
    sCodigoContato = CodContato.Text

    'Limpa os MaskedEdit e TextBoxes
    Call Limpa_Tela(Me)
    
    Ativo.Value = MARCADO
    
    'Mantém o código de Contato na Tela
    CodContato.Text = sCodigoContato
    
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

    'Limpa o grid de Categorias
    Call Grid_Limpa(objGridCategoria)
    
    TipoFrete.ListIndex = -1
        
    Mensagem.Text = "" 'Inserido por Wagner 26/09/05

    Call gobjTabEnd.Limpa_Tela
    
    Limpa_Tela_FiliaisContatos = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_FiliaisContatos:
    
    Limpa_Tela_FiliaisContatos = gErr
    
    Select Case gErr
        
        Case 180930
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180931)

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
    Set objEventoContato = New AdmEvento
    Set objEventoFilialContato = New AdmEvento
    Set objEventoMensagem = New AdmEvento
        
    iFrameAtual1 = 1
    iFrameAtual2 = 0

    OpcaoEndereco(0) = True
    OpcaoEndereco(1) = False
    OpcaoEndereco(2) = False
    
    Ativo.Value = MARCADO

    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then gError 180932

    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        Set objEventoContaContabil = New AdmEvento
        
        'Inicializa propriedade Mask de ContaContabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 180935
    
        ContaContabil.Mask = sMascaraConta

    Else
       
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then gError 180935
    
        ContaContabil.Mask = sMascaraConta
        
        'Desabilita conta contábil
        ContaContabil.Enabled = False
        ContaContabilLabel.Enabled = False
        
    End If

    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then gError 180936

    'Preenche a ComboBox Regiao com os objetos da colecao colCodigoNome
    For Each objCodigoNome In colCodigoNome
        Regiao.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoNome.iCodigo
    Next

    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 180937

    'Preenche a ComboBox Cobrador com os objetos da colecao colCodigoNome
    For Each objCobrador In ColCobrador
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If
    Next

    'Lê códigos e nomes reduzidos da tabela Transportadoras
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 180938

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
    If lErro <> SUCESSO Then gError 180939

    'Carrega a combobox de Categoria Produto
    lErro = Carrega_ComboCategoriaContato()
    If lErro <> SUCESSO Then gError 180940
    
    'Preenche Combo Mensagem
    Set colCodigoNome = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela Mensagem
    lErro = CF("Cod_Nomes_Le", "Mensagens", "Codigo", "Descricao", STRING_NFISCAL_MENSAGEM, colCodigoNome)
    If lErro <> SUCESSO Then gError 180941

    'Preenche a ComboBox Mensagem com os objetos da colecao colCodigoDescricao
    For Each objCodigoNome In colCodigoNome
        Mensagem.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Mensagem.ItemData(Mensagem.NewIndex) = objCodigoNome.iCodigo
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 180932 To 180941

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180942)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFilialContato As ClassFilialContato) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objFilialContato Is Nothing) Then

        'Se foi passado código de Contato
        If objFilialContato.lCodContato <> 0 Then

            'Se foi passado código de Filial
            If objFilialContato.iCodFilial <> 0 Then

                'Tenta ler FilialContato
                lErro = CF("FilialContato_Le", objFilialContato)
                If lErro <> SUCESSO And lErro <> 52591 Then gError 180943

                'Se a Filial não existir
                If lErro <> SUCESSO Then
                    
                    'Limpa a Tela
                    lErro = Limpa_Tela_FiliaisContatos()
                    If lErro <> SUCESSO Then gError 180944
                    
                    CodContato.Text = CStr(objFilialContato.lCodContato)
                    CodFilial.Text = CStr(objFilialContato.iCodFilial)
                    
                    
                Else  'Filial existe, então exibe seus dados

                    lErro = Exibe_Dados_FilialContato(objFilialContato)
                    If lErro <> SUCESSO Then gError 180945

                End If

            Else 'Apenas o código do Contato foi passado
                
                'Limpa Tela
                lErro = Limpa_Tela_FiliaisContatos()
                If lErro <> SUCESSO Then gError 180946
                
                CodContato.Text = CStr(objFilialContato.lCodContato)

                'Inserido por Leo em 10/01/02
                If Len(Trim(objFilialContato.sNomeReduzidoCli)) > 0 Then
                    
                    For iIndice = 0 To 4
                        ContatoLabel(iIndice).Caption = objFilialContato.sNomeReduzidoCli
                    Next

                End If
                
                If Len(Trim(objFilialContato.sNome)) > 0 Then
                    Nome.Text = objFilialContato.sNome
                End If
                'Leo até aqui

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 180943, 180944, 180945, 180946

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180947)

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

Private Function Carrega_ComboCategoriaContato() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaContato As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaContato

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos", colCategorias)
    If lErro <> SUCESSO Then gError 180948

    For Each objCategoriaContato In colCategorias

        'Insere na combo CategoriaContato
        ComboCategoriaContato.AddItem objCategoriaContato.sCategoria

    Next

    Carrega_ComboCategoriaContato = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaContato:

    Carrega_ComboCategoriaContato = gErr

    Select Case gErr

        Case 180948

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180949)

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
    
End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual1 Then

        If TabStrip_PodeTrocarTab(iFrameAtual1, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual1).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual1 = Opcao.SelectedItem.Index
               
    End If

End Sub

Function Exibe_Dados_FilialContato(objFilialContato As ClassFilialContato) As Long
'Exibe os dados da filial do Contato na tela

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim colCategoriaItem As New Collection
Dim bCancel As Boolean
Dim objContato As New ClassContatos
Dim objCliente As New ClassCliente

On Error GoTo Erro_Exibe_Dados_FilialContato

    objContato.lCodigo = objFilialContato.lCodContato
    
    objCliente.lEndereco = objFilialContato.lEndereco
    objCliente.lEnderecoCobranca = objFilialContato.lEnderecoCobranca
    objCliente.lEnderecoEntrega = objFilialContato.lEnderecoEntrega

    'Lê Contato a partir do Nome Reduzido
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180748 Then gError 180953

    'Contato não cadastrado
    If lErro = 180748 Then gError 180954

    For iIndice = 0 To 3
        ContatoLabel(iIndice).Caption = objContato.sNomeReduzido
    Next

    'IDENTIFICACAO :

    CodContato.Text = CStr(objFilialContato.lCodContato)
    CodFilial.Text = CStr(objFilialContato.iCodFilial)
    Nome.Text = objFilialContato.sNome
        
    'Lê as Categoria da Filial do Contato
    lErro = CF("FilialContato_Le_Categorias", objFilialContato, colCategoriaItem)
    If lErro <> SUCESSO Then gError 180955
    
    Ativo.Value = objFilialContato.iAtivo
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

    FilialLabel(2).Caption = objFilialContato.sNome
    
    RG.Text = objFilialContato.sRG
    CGC.Text = objFilialContato.sCgc
    Call CGC_Validate(bSGECancelDummy)
    InscricaoEstadual.Text = objFilialContato.sInscricaoEstadual
    InscricaoMunicipal.Text = objFilialContato.sInscricaoMunicipal
    Observacao.Text = objFilialContato.sObservacao
    
    InscricaoSuframa.PromptInclude = False
    InscricaoSuframa.Text = objFilialContato.sInscricaoSuframa
    InscricaoSuframa.PromptInclude = True
    
    'ENDERECOS :

    FilialLabel(0).Caption = objFilialContato.sNome

    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
    If lErro <> SUCESSO Then gError 180956

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then gError 180956

    'VENDAS :

    FilialLabel(1).Caption = objFilialContato.sNome

    If objFilialContato.iVendedor = 0 Then
        Vendedor.Text = ""
    Else
        Vendedor.Text = CStr(objFilialContato.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If

    ContaContabil.PromptInclude = False

    If Len(Trim(objFilialContato.sContaContabil)) = 0 Then
        ContaContabil.Text = objFilialContato.sContaContabil
    Else
        lErro = Mascara_RetornaContaEnxuta(objFilialContato.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 180957
        ContaContabil.Text = sContaEnxuta
    End If

    ContaContabil.PromptInclude = True

    If objFilialContato.dComissaoVendas = 0# Then
        ComissaoVendas.Text = ""
    Else
        ComissaoVendas.Text = CStr(100 * objFilialContato.dComissaoVendas)
    End If

    If objFilialContato.iRegiao = 0 Then
        Regiao.Text = ""
    Else
        Regiao.Text = CStr(objFilialContato.iRegiao)
        Call Regiao_Validate(bCancel)
    End If

    If objFilialContato.iCodCobrador = 0 Then
        Cobrador.Text = ""
    Else
        Cobrador.Text = CStr(objFilialContato.iCodCobrador)
        Call Cobrador_Validate(bCancel)
    End If

    If objFilialContato.iFreqVisitas = 0 Then
        FreqVisitas.Text = ""
    Else
        FreqVisitas.Text = CStr(objFilialContato.iFreqVisitas)
    End If

    DataUltVisita.PromptInclude = False
    If (objFilialContato.dtDataUltVisita) = DATA_NULA Then
        DataUltVisita.Text = "  /  /  "
    Else
        DataUltVisita.Text = Format(objFilialContato.dtDataUltVisita, "dd/mm/yy")
    End If

    DataUltVisita.PromptInclude = True
    
    'Tipo de Frete
    If objFilialContato.iTipoFrete <> -1 Then
        For iIndice = 0 To TipoFrete.ListCount - 1
            If TipoFrete.ItemData(iIndice) = objFilialContato.iTipoFrete Then
                TipoFrete.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoFrete.ListIndex = -1
    End If
    
    If objFilialContato.iCodTransportadora = 0 Then
        Transportadora.Text = ""
    Else
        Transportadora.Text = CStr(objFilialContato.iCodTransportadora)
        Call Transportadora_Validate(bCancel)
    End If
    
    Guia.Text = objFilialContato.sGuia
    
    If objFilialContato.iCodTranspRedesp = 0 Then
        TranspRedespacho.Text = ""
    Else
        TranspRedespacho.Text = CStr(objFilialContato.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If
    
    RedespachoCli.Value = objFilialContato.iDetPagFrete

    If objFilialContato.iCodMensagem = 0 Then
        Mensagem.Text = ""
    Else
        Mensagem.Text = CStr(objFilialContato.iCodMensagem)
        Call Mensagem_Validate(bSGECancelDummy)
    End If

    iAlterado = 0

    Exibe_Dados_FilialContato = SUCESSO

Exit Function

Erro_Exibe_Dados_FilialContato:

    Exibe_Dados_FilialContato = gErr

    Select Case gErr

        Case 180953, 180956, 180955

        Case 180954
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)
        
        Case 180957
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objFilialContato.sContaContabil)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180958)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFilialContato As New ClassFilialContato
Dim colEndereco As New Collection
Dim colCategoriaItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Codigo do Contato
    If Len(Trim(CodContato.Text)) = 0 Then gError 180959

    'Verifica se foi preenchido Codigo da Filial
    If Len(Trim(CodFilial.Text)) = 0 Then gError 180960

    'Verifica se foi preenchido o Nome da Filial
    If Len(Trim(Nome.Text)) = 0 Then gError 180961

'    'Verifica se foi preenchido o Estado
'    If Len(Trim(Endereco(0).Text)) <> 0 Then
'        If Len(Trim(Estado(0).Text)) = 0 Then gError 180962
'    End If
'
'    If Len(Trim(Endereco(1).Text)) <> 0 Then
'        If Len(Trim(Estado(1).Text)) = 0 Then gError 180963
'    End If
'
'    If Len(Trim(Endereco(2).Text)) <> 0 Then
'        If Len(Trim(Estado(2).Text)) = 0 Then gError 180964
'    End If

    'Verifica se para cada Categoria existente no grid de Categorias existe um Valor correspondente
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 And Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL))) = 0 Then gError 180970
    Next

    'Lê os Enderecos e coloca em colEndereco
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEndereco)
    If lErro <> SUCESSO Then gError 180965

    'Lê os dados da tela relativos à Filial Contato
    lErro = Le_Dados_FilialContato(objFilialContato)
    If lErro <> SUCESSO Then gError 180966
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objFilialContato.sCgc)) > 0 Then
        'Verifica se tem outro Contato com o mesmo CGC e dá aviso
        lErro = CF("FilialCliente_Testa_CGC", 0, 0, objFilialContato.sCgc)
        If lErro <> SUCESSO Then gError 180967
    End If
    
    lErro = Trata_Alteracao(objFilialContato, objFilialContato.lCodContato, objFilialContato.iCodFilial)
    If lErro <> SUCESSO Then gError 180968
    
    'Grava FilialContato no BD
    lErro = CF("FilialContatos_Grava", objFilialContato, colEndereco)
    If lErro <> SUCESSO Then gError 180969

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 180959
            Call Rotina_Erro(vbOKOnly, "ERRO_CODContato_NAO_PREENCHIDO", gErr)

        Case 180960
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 180961
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEFILIAL_NAO_PREENCHIDO", gErr)

        Case 180962, 180963, 180964
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case 180965, 180966, 180967, 180968, 180969

        Case 180970
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180971)

    End Select

    Exit Function

End Function

Private Function Le_Dados_FilialContato(objFilialContato As ClassFilialContato) As Long
'Lê os dados que estão na tela de FiliaisContatos e coloca-os em objFilialContato

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sConta As String
Dim iContaPreenchida As Integer
Dim objVendedor As New ClassVendedor
Dim objFilialCliCategoria As New ClassFilialConCategoria
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objCobrador As New ClassCobrador
Dim objTransportadora As New ClassTransportadora
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Le_Dados_FilialContato

    'IDENTIFICACAO :
    
    objFilialContato.iAtivo = Ativo.Value
    
    If Len(Trim(CodContato.Text)) > 0 Then objFilialContato.lCodContato = CLng(CodContato.Text)

    If Len(Trim(CodFilial.Text)) > 0 Then objFilialContato.iCodFilial = CodFilial.Text
    
    objFilialContato.sRG = Trim(RG.Text)
    objFilialContato.sNome = Trim(Nome.Text)
    objFilialContato.sCgc = Trim(CGC.Text)
    objFilialContato.sInscricaoEstadual = Trim(InscricaoEstadual.Text)
    objFilialContato.sInscricaoMunicipal = Trim(InscricaoMunicipal.Text)
    objFilialContato.sInscricaoSuframa = Trim(InscricaoSuframa.ClipText)
    objFilialContato.sObservacao = Trim(Observacao.Text)

    'Preenche uma coleção com todas as linhas "existentes" do grid de categorias
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL))) <> 0 Then

            Set objFilialCliCategoria = New ClassFilialConCategoria

            objFilialCliCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL)
            objFilialCliCategoria.sItem = GridCategoria.TextMatrix(iIndice, GRID_VALOR_COL)

            'Verifica se já existe a categoria na coleção
            If objFilialContato.colCategoriaItem.Count > 0 Then

                For iIndice1 = 1 To objFilialContato.colCategoriaItem.Count

                    If objFilialCliCategoria.sCategoria = objFilialContato.colCategoriaItem.Item(iIndice1).sCategoria Then gError 180985

                Next

            End If

            objFilialContato.colCategoriaItem.Add objFilialCliCategoria

        End If

    Next

   'VENDAS :

    If Len(Trim(Vendedor.Text)) > 0 Then

        objVendedor.sNomeReduzido = Vendedor.Text

        'Tenta ler Vendedor a partir do NomeReduzido
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 180974

        'Não encontrou o Vendedor
        If lErro <> SUCESSO Then gError 180975

        objFilialContato.iVendedor = objVendedor.iCodigo

    End If

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 180976

    If iContaPreenchida = CONTA_VAZIA Then
        objFilialContato.sContaContabil = ""
    Else
        objFilialContato.sContaContabil = sConta
    End If

    If Len(Trim(ComissaoVendas.Text)) > 0 Then objFilialContato.dComissaoVendas = CDbl(ComissaoVendas.Text) / 100

    If Len(Trim(Regiao.Text)) > 0 Then objFilialContato.iRegiao = Codigo_Extrai(Regiao.Text)

    'Verifica se região venda existe
    If objFilialContato.iRegiao <> 0 Then
        objRegiaoVenda.iCodigo = objFilialContato.iRegiao
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 180977

        'Não encontrou a região venda ==> erro
        If lErro = 16137 Then gError 180978

    End If

    If Len(Trim(Cobrador.Text)) > 0 Then objFilialContato.iCodCobrador = Codigo_Extrai(Cobrador.Text)
    
    'Verifica se cobrador existe
    If objFilialContato.iCodCobrador <> 0 Then
        objCobrador.iCodigo = objFilialContato.iCodCobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 180979

        'Não encontrou o cobrador ==> erro
        If lErro = 19294 Then gError 180980

    End If

    If Len(Trim(FreqVisitas.Text)) > 0 Then objFilialContato.iFreqVisitas = CInt(FreqVisitas.Text)

    If Trim(DataUltVisita.Text) = "/  /" Then
        objFilialContato.dtDataUltVisita = DATA_NULA
    Else
        objFilialContato.dtDataUltVisita = CDate(DataUltVisita.Text)
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then objFilialContato.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    
    'Tipo de Frete
    If TipoFrete.ListIndex <> -1 Then
        If TipoFrete.ItemData(TipoFrete.ListIndex) = TIPO_CIF Then
            objFilialContato.iTipoFrete = TIPO_CIF
        Else
            objFilialContato.iTipoFrete = TIPO_FOB
        End If
    Else
        objFilialContato.iTipoFrete = -1
    End If
    
    'Verifica se transportadora existe
    If objFilialContato.iCodTransportadora <> 0 Then
        
        objTransportadora.iCodigo = objFilialContato.iCodTransportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180981

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 180982

    End If

    objFilialContato.sGuia = Trim(Guia.Text)
    
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objFilialContato.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)

    'Verifica se transportadora existe
    If objFilialContato.iCodTranspRedesp <> 0 Then
        objTransportadora.iCodigo = objFilialContato.iCodTranspRedesp
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 180981

        'Não achou a transportadora ==> erro
        If lErro = 19250 Then gError 180982

    End If
    
    objFilialContato.iDetPagFrete = RedespachoCli.Value

    If Len(Trim(Mensagem.Text)) > 0 Then objFilialContato.iCodMensagem = Codigo_Extrai(Mensagem.Text)

    'Verifica se mensagem existe
    If objFilialContato.iCodMensagem <> 0 Then
        objMensagem.iCodigo = objFilialContato.iCodMensagem
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 180983

        'Não encontrou a mensagem ==> erro
        If lErro = 19234 Then gError 180984

    End If

    Le_Dados_FilialContato = SUCESSO

    Exit Function

Erro_Le_Dados_FilialContato:

    Le_Dados_FilialContato = gErr

    Select Case gErr

        Case 180974, 180976, 180977, 180979, 180981, 180983
            'Erro tratado na rotina chamada

        Case 180975
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 180978
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case 180980
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)

        Case 180982
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
        
        Case 180984
            Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", gErr, objMensagem.iCodigo)

        Case 180985
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_REPETIDA_NO_GRID", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180986)

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
    
    'Chama a tela que lista os Contatos
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
    If lErro <> SUCESSO Then gError 180991
    
    'Se a conta estiver preenchida passa para objPlanoConta
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta
    
    'Chama a tela que lista as contas
    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

    Case 180991

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180992)

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

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 180993

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 180993
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180994)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA  "
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFilialContato As New ClassFilialContato

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FiliaisContatos"

    'Lê os dados da Tela FilialContato
    lErro = Le_Dados_FilialContato(objFilialContato)
    If lErro <> SUCESSO Then gError 180995

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodContato", objFilialContato.lCodContato, 0, "CodContato"
    colCampoValor.Add "CodFilial", objFilialContato.iCodFilial, 0, "CodFilial"
    colCampoValor.Add "Nome", objFilialContato.sNome, STRING_FILIAL_CLIENTE_NOME, "Nome"
    colCampoValor.Add "RG", objFilialContato.sRG, STRING_RG, "RG"
    colCampoValor.Add "CGC", objFilialContato.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objFilialContato.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "InscricaoMunicipal", objFilialContato.sInscricaoMunicipal, STRING_INSCR_MUN, "InscricaoMunicipal"
    colCampoValor.Add "InscricaoSuframa", objFilialContato.sInscricaoSuframa, STRING_INSCR_SUF, "InscricaoSuframa"
    colCampoValor.Add "Endereco", objFilialContato.lEndereco, 0, "Endereco"
    colCampoValor.Add "EnderecoEntrega", objFilialContato.lEnderecoEntrega, 0, "EnderecoEntrega"
    colCampoValor.Add "EnderecoCobranca", objFilialContato.lEnderecoCobranca, 0, "EnderecoCobranca"
    colCampoValor.Add "CodTransportadora", objFilialContato.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objFilialContato.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objFilialContato.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "Guia", objFilialContato.sGuia, STRING_CLIENTE_GUIA, "Guia"
    colCampoValor.Add "Observacao", objFilialContato.sObservacao, STRING_CLIENTE_OBSERVACAO, "Observacao"
    colCampoValor.Add "ContaContabil", objFilialContato.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "Vendedor", objFilialContato.iVendedor, 0, "Vendedor"
    colCampoValor.Add "ComissaoVendas", objFilialContato.dComissaoVendas, 0, "ComissaoVendas"
    colCampoValor.Add "Regiao", objFilialContato.iRegiao, 0, "Regiao"
    colCampoValor.Add "FreqVisitas", objFilialContato.iFreqVisitas, 0, "FreqVisitas"
    colCampoValor.Add "DataUltVisita", objFilialContato.dtDataUltVisita, 0, "DataUltVisita"
    colCampoValor.Add "CodCobrador", objFilialContato.iCodCobrador, 0, "CodCobrador"
    colCampoValor.Add "TipoFrete", objFilialContato.iTipoFrete, 0, "TipoFrete"
    colCampoValor.Add "Ativo", objFilialContato.iAtivo, 0, "Ativo"
    colCampoValor.Add "CodMensagem", objFilialContato.iCodMensagem, 0, "CodMensagem" 'Inserido por Wagner '18/11/2005
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 180995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180996)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objFilialContato As New ClassFilialContato

On Error GoTo Erro_Tela_Preenche

    objFilialContato.lCodContato = colCampoValor.Item("CodContato").vValor

    'Se código de Contato estiver preenchido,
    If objFilialContato.lCodContato <> 0 Then

        'Passa os dados da coleção para o objeto
        objFilialContato.iCodFilial = colCampoValor.Item("CodFilial").vValor
        objFilialContato.sNome = colCampoValor.Item("Nome").vValor
        objFilialContato.sCgc = colCampoValor.Item("CGC").vValor
        objFilialContato.sRG = colCampoValor.Item("RG").vValor
        objFilialContato.sInscricaoEstadual = colCampoValor.Item("InscricaoEstadual").vValor
        objFilialContato.sInscricaoMunicipal = colCampoValor.Item("InscricaoMunicipal").vValor
        objFilialContato.sInscricaoSuframa = colCampoValor.Item("InscricaoSuframa").vValor
        objFilialContato.lEndereco = colCampoValor.Item("Endereco").vValor
        objFilialContato.lEnderecoEntrega = colCampoValor.Item("EnderecoEntrega").vValor
        objFilialContato.lEnderecoCobranca = colCampoValor.Item("EnderecoCobranca").vValor
        objFilialContato.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
        objFilialContato.iCodTranspRedesp = colCampoValor.Item("CodTranspRedesp").vValor
        objFilialContato.iDetPagFrete = colCampoValor.Item("DetPagFrete").vValor
        objFilialContato.sGuia = colCampoValor.Item("Guia").vValor
        objFilialContato.sObservacao = colCampoValor.Item("Observacao").vValor
        objFilialContato.sContaContabil = colCampoValor.Item("ContaContabil").vValor
        objFilialContato.iVendedor = colCampoValor.Item("Vendedor").vValor
        objFilialContato.dComissaoVendas = colCampoValor.Item("ComissaoVendas").vValor
        objFilialContato.iRegiao = colCampoValor.Item("Regiao").vValor
        objFilialContato.iFreqVisitas = colCampoValor.Item("Freqvisitas").vValor
        objFilialContato.dtDataUltVisita = colCampoValor.Item("DataUltVisita").vValor
        objFilialContato.iCodCobrador = colCampoValor.Item("CodCobrador").vValor
        objFilialContato.iTipoFrete = colCampoValor.Item("TipoFrete").vValor
        objFilialContato.iAtivo = colCampoValor.Item("Ativo").vValor
        objFilialContato.iCodMensagem = colCampoValor.Item("CodMensagem").vValor 'Inserido por Wagner
               
        'Exibe os dados da FilialContato
        lErro = Exibe_Dados_FilialContato(objFilialContato)
        If lErro <> SUCESSO Then gError 180997

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 180997

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180998)

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
    Set objEventoContato = Nothing
    Set objEventoFilialContato = Nothing
    
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
                If lErro <> SUCESSO Then gError 180999

            Case GRID_VALOR_COL

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 181000

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 181001

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 180999, 181000

        Case 181001
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181002)

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

    iIndice = GridCategoria.Row
    If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) <> ComboCategoriaContato.Text Then

        If Len(Trim(ComboCategoriaContato.Text)) > 0 Then

            'Preenche o objeto com a Categoria
            objCategoriaContato.sCategoria = ComboCategoriaContato.Text

            'Lê Categoria De Contato no BD
            lErro = CF("CategoriaCliente_Le", objCategoriaContato)
            If lErro <> SUCESSO And lErro <> 28847 Then gError 181003

            'Categoria não está cadastrada
            If lErro = 28847 Then gError 181004

            'Verifica se já existe a categoria no Grid
            If GridCategoria.Row > 0 Then

                iIndice1 = GridCategoria.Row

                For iIndice = 1 To objGridCategoria.iLinhasExistentes

                    If iIndice <> iIndice1 Then

                        If GridCategoria.TextMatrix(iIndice, GRID_CATEGORIA_COL) = ComboCategoriaContato.Text Then gError 181005

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
    If lErro <> SUCESSO Then gError 181006

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 181003
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 181004 'Categoria não está cadastrada
            
            'Perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CRIAR_CATEGORIACLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a Tela "CategoriaContato"
                Call Chama_Tela("CategoriaCliente", objCategoriaContato)

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 181005
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 181006
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181007)
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
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)) = 0 Then gError 181008

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaContatoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaContatoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, GRID_CATEGORIA_COL)
            objCategoriaContatoItem.sItem = ComboCategoriaContatoItem.Text

            'Lê Item De Categoria De Contato no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaContatoItem)
            If lErro <> SUCESSO And lErro <> 28991 Then gError 181009

            'Item da Categoria não está cadastrado
            If lErro = 28991 Then gError 181010

        End If

    End If

    If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then

        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 181011

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 181008
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 181009, 181011
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 181010 'Item da Categoria não está cadastrado

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

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181012)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
    Caption = "Filiais de Clientes Futuros"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FiliaisContatos"
    
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
        ElseIf Me.ActiveControl Is CodContato Then
            Call LabelCodContato_Click
        ElseIf Me.ActiveControl Is CodFilial Then
            Call LabelCodFilial_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is Mensagem Then
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 181013

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 181014
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 181015
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 181016

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 181013, 181014

        Case 181015  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 181016
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181017)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 181022

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objMensagem.iCodigo = iCodigo

        'Tenta ler Mensagem com esse código no BD
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 181023

        If lErro <> SUCESSO Then gError 181024 'Não encontrou Mensagem no BD

        'Encontrou Mensagem no BD, coloca no Text da Combo
        Mensagem.Text = CStr(objMensagem.iCodigo) & SEPARADOR & objMensagem.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 181025

    Exit Sub

Erro_Mensagem_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 181022, 181023

    Case 181024  'Não encontrou Mensagem no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MENSAGEM")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Mensagens", objMensagem)

        End If

    Case 181025
        Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_ENCONTRADA", gErr, Mensagem.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181026)

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

