VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'??? tab index
'??? validar CGC, fazer com que formate. ver como foi feito em ctfiliaisclientes.cls

Dim m_objUserControl As Object
Public gobjTabEnd As New ClassTabEndereco

Public gobjInfoUsu As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Public iAlterado As Integer
Dim iEmEvento As Integer
Dim iTipoAlterado As Integer

Private WithEvents objEventoTipoVendedor As AdmEvento
Attribute objEventoTipoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoBanco As AdmEvento
Attribute objEventoBanco.VB_VarHelpID = -1
Private WithEvents objEventoSuperior As AdmEvento
Attribute objEventoSuperior.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Comissao = 2
Private Const TAB_Endereco = 3

'--- inicio dos properties get dos controles da tela
Public Property Get RG() As Object
     Set RG = objUserControl.Controls("RG")
End Property

Public Property Get CGC0() As Object
     Set CGC0 = objUserControl.Controls("CGC0")
End Property

Public Property Get CodUsuario() As Object
     Set CodUsuario = objUserControl.Controls("CodUsuario")
End Property

Public Property Get Cargo() As Object
     Set Cargo = objUserControl.Controls("Cargo")
End Property

Public Property Get Superior() As Object
     Set Superior = objUserControl.Controls("Superior")
End Property

Public Property Get FrameInfoVinculo() As Object
     Set FrameInfoVinculo = objUserControl.Controls("FrameInfoVinculo")
End Property

Public Property Get RazaoSocial() As Object
     Set RazaoSocial = objUserControl.Controls("RazaoSocial")
End Property

Public Property Get CGC() As Object
     Set CGC = objUserControl.Controls("CGC")
End Property

Public Property Get InscricaoEstadual() As Object
     Set InscricaoEstadual = objUserControl.Controls("InscricaoEstadual")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Regiao() As Object
     Set Regiao = objUserControl.Controls("Regiao")
End Property

Public Property Get VendedoresList() As Object
     Set VendedoresList = objUserControl.Controls("VendedoresList")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get Nome() As Object
     Set Nome = objUserControl.Controls("Nome")
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objUserControl.Controls("NomeReduzido")
End Property

Public Property Get Matricula() As Object
     Set Matricula = objUserControl.Controls("Matricula")
End Property

Public Property Get TipoLabel() As Object
     Set TipoLabel = objUserControl.Controls("TipoLabel")
End Property

Public Property Get RegiaoVendaLabel() As Object
     Set RegiaoVendaLabel = objUserControl.Controls("RegiaoVendaLabel")
End Property

Public Property Get ComissaoSobreTotal() As Object
     Set ComissaoSobreTotal = objUserControl.Controls("ComissaoSobreTotal")
End Property

Public Property Get ComissaoVenda() As Object
     Set ComissaoVenda = objUserControl.Controls("ComissaoVenda")
End Property

Public Property Get ComissaoFrete() As Object
     Set ComissaoFrete = objUserControl.Controls("ComissaoFrete")
End Property

Public Property Get ComissaoSeguro() As Object
     Set ComissaoSeguro = objUserControl.Controls("ComissaoSeguro")
End Property

Public Property Get ComissaoICM() As Object
     Set ComissaoICM = objUserControl.Controls("ComissaoICM")
End Property

Public Property Get ComissaoIPI() As Object
     Set ComissaoIPI = objUserControl.Controls("ComissaoIPI")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get SaldoComissao() As Object
     Set SaldoComissao = objUserControl.Controls("SaldoComissao")
End Property

Public Property Get DataUltVenda() As Object
     Set DataUltVenda = objUserControl.Controls("DataUltVenda")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get Agencia() As Object
     Set Agencia = objUserControl.Controls("Agencia")
End Property

Public Property Get Banco() As Object
     Set Banco = objUserControl.Controls("Banco")
End Property

Public Property Get Ativo() As Object
     Set Ativo = objUserControl.Controls("Ativo")
End Property

Public Property Get LabelBanco() As Object
     Set LabelBanco = objUserControl.Controls("LabelBanco")
End Property

Public Property Get VendedorLabel() As Object
     Set VendedorLabel = objUserControl.Controls("VendedorLabel")
End Property

Public Property Get PercComissao() As Object
     Set PercComissao = objUserControl.Controls("PercComissao")
End Property

Public Property Get PercComissaoEmissao() As Object
     Set PercComissaoEmissao = objUserControl.Controls("PercComissaoEmissao")
End Property

Public Property Get PercComissaoBaixa() As Object
     Set PercComissaoBaixa = objUserControl.Controls("PercComissaoBaixa")
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

'Incluido por Mauricio Maciel 11/04/03
Public Property Get Vinculo() As Object
     Set Vinculo = objUserControl.Controls("Vinculo")
End Property

Public Property Get TabEnd() As Object
     Set TabEnd = objUserControl.Controls("TabEnd")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Ativo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Banco_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do próximo Vendedor
    lErro = CF("Vendedor_Automatico", sCodigo)
    If lErro <> SUCESSO Then Error 57560

    Codigo.Text = sCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57560
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158402)
    
    End Select

    Exit Sub

End Sub

Public Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16694

    objVendedor.iCodigo = CInt(Codigo.Text)

    'Lê o Vendedor
    lErro = CF("Vendedor_Le", objVendedor)
    If lErro <> SUCESSO And lErro <> 12582 Then Error 16696
    
    'Se não encontrou o Vendedor --> Erro
    If lErro = 12582 Then Error 16695

    'Pedido de confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_VENDEDOR", objVendedor.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Vendedor
        lErro = CF("Vendedor_Exclui", objVendedor)
        If lErro <> SUCESSO Then Error 16697

        Call VendedoresList_Exclui(objVendedor.iCodigo)
        
        Call Limpa_Tela_Vendedor

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
        
        Case 16696, 16697

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158403)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Chama a função de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16655
    
    'Limpa a tela
    Call Limpa_Tela_Vendedor
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 16655
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158404)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80460

    Call Limpa_Tela_Vendedor

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80460

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158405)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Critica se é inteiro
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 16650

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16650

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158406)

    End Select

    Exit Sub
    
End Sub

Public Sub ComissaoFrete_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1
        
        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal = vbChecked
        Else
            ComissaoSobreTotal = vbUnchecked
        End If

        iEmEvento = 0

    End If
    
    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ComissaoICM_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal = vbChecked
        Else
            ComissaoSobreTotal = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoIPI_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal = vbChecked
        Else
            ComissaoSobreTotal = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ComissaoSeguro_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal = vbChecked
        Else
            ComissaoSobreTotal = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoSobreTotal_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1
        
        'se tudo foi selecionado então marca todas as comissões
        If ComissaoSobreTotal.Value = vbChecked Then
            ComissaoFrete.Value = vbChecked
            ComissaoICM.Value = vbChecked
            ComissaoIPI.Value = vbChecked
            ComissaoSeguro.Value = vbChecked
        Else
            'desmarca todas as comissões
            ComissaoFrete.Value = vbUnchecked
            ComissaoICM.Value = vbUnchecked
            ComissaoIPI.Value = vbUnchecked
            ComissaoSeguro.Value = vbUnchecked
        End If

        iEmEvento = 0

    End If

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Function Incidencias_Selecionadas() As CheckBoxConstants
'retorna vbChecked se a base da comissao inclui todas opcoes (frete, seguro,...)

    Incidencias_Selecionadas = ComissaoFrete.Value * ComissaoICM.Value * ComissaoIPI.Value * ComissaoSeguro.Value

End Function

Public Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objVendedor As New ClassVendedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult
Dim colEnderecos As New Collection
Dim objVendedor1 As New ClassVendedor

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then gError 16657

    'Verifica preenchimento do nome
    If Len(Trim(Nome.Text)) = 0 Then gError 16658
    
    'Verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 16659
    
    If Len(Trim(PercComissaoEmissao.Text)) = 0 Then gError 16720
    If CDbl(PercComissaoEmissao.Text) + CDbl(left(PercComissaoBaixa.Caption, Len(PercComissaoBaixa.Caption) - 1)) <> 100 Then gError 25226
                                  
    objVendedor.iCargo = Codigo_Extrai(Cargo.Text)
    If Len(Trim(Superior.Text)) > 0 Then objVendedor1.sNomeReduzido = Superior.Text

    'Verifica se vendedor existe
    If objVendedor1.sNomeReduzido <> "" Then
        
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor1)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 195888

        'Não encontrou o superior ==> erro
        If lErro = 25008 Then gError 195889
        
        'o superior nao pode ser o proprio vendedor
        If objVendedor.iCodigo = objVendedor1.iCodigo Then gError 195890

        objVendedor.iSuperior = objVendedor1.iCodigo

    End If
    
    'Preenche objVendedor
    lErro = Le_Dados_Vendedor(objVendedor)
    If lErro <> SUCESSO Then gError 16662
    
    'Preenche objEndereco
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then gError 16660
    
    Set objEndereco = colEnderecos.Item(1)
    
    lErro = Trata_Alteracao(objVendedor, objVendedor.iCodigo)
    If lErro <> SUCESSO Then gError 16715
    
    lErro = CF2(Me, "VND_Grava_Registro", objVendedor)
    If lErro <> SUCESSO Then gError 195869
    
    'Chama função de gravação
    lErro = CF("Vendedor_Grava", objVendedor, objEndereco)
    If lErro <> SUCESSO Then gError 16663

    'Remove o Vendedor da ListBox VendedoresList
    Call VendedoresList_Exclui(objVendedor.iCodigo)
    
    'Insere o Vendedor na ListBox VendedoresList
    Call VendedoresList_Adiciona(objVendedor)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 16657
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 16658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 16659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
        
        Case 16660, 16662, 16663, 16715, 195869
        
        Case 16720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTAGEM_EMISSAO_NAO_PREENCHIDA", gErr)

        Case 25226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_EMISSAO_MAIS_BAIXA", gErr)

        Case 195888
        
        Case 195889
            Call Rotina_Erro(vbOKOnly, "ERRO_SUPERIOR_NAO_CADASTRADO", gErr)
        
        Case 195890
            Call Rotina_Erro(vbOKOnly, "ERRO_SUPERIOR_IGUAL_VENDEDOR", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158408)

    End Select
    
    Exit Function
        
End Function

Public Sub VendedoresList_Adiciona(objVendedor As ClassVendedor)

    VendedoresList.AddItem objVendedor.sNomeReduzido
    VendedoresList.ItemData(VendedoresList.NewIndex) = objVendedor.iCodigo

End Sub

Public Sub VendedoresList_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To VendedoresList.ListCount - 1

        If VendedoresList.ItemData(iIndice) = iCodigo Then

            VendedoresList.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Public Sub Fax_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()
        
Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim colCodigoNome As AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim colCodigo As Collection
Dim vCodigo As Variant
Dim objTela As Object

On Error GoTo Erro_Form_Load
    
    Set objEventoTipoVendedor = New AdmEvento
    Set objEventoBanco = New AdmEvento
    Set objEventoSuperior = New AdmEvento
    
    iFrameAtual = 1
    iAlterado = 0
    Ativo.Value = MARCADO
    
    Set colCodigoNome = New AdmColCodigoNome

    'Leitura dos códigos e nome reduzidos dos Vendedores
    lErro = CF("Cod_Nomes_Le", "Vendedores", "Codigo", "NomeReduzido", STRING_VENDEDOR_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 16612

    'Preenche ListBox VendedoresList com nome reduzido dos vendedores
    For Each objCodigoNome In colCodigoNome
        VendedoresList.AddItem objCodigoNome.sNome
        VendedoresList.ItemData(VendedoresList.NewIndex) = objCodigoNome.iCodigo
    Next

    Set colCodigoNome = New AdmColCodigoNome

    'Leitura dos códigos e descrições dos Tipos De Vendedor
    lErro = CF("Cod_Nomes_Le", "TiposDeVendedor", "Codigo", "Descricao", STRING_TIPO_DE_VENDEDOR_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then Error 16613

    'Preenche a ComboBox Tipo com código e descrição dos Tipos De Vendedor
    For Each objCodigoNome In colCodigoNome
        Tipo.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Tipo.ItemData(Tipo.NewIndex) = objCodigoNome.iCodigo
    Next

    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CARGO_VENDEDOR, Cargo)
    If lErro <> SUCESSO Then Error 16613

    'Carrega a combobox todos os usuários
    lErro = Carrega_Usuarios()
    If lErro <> SUCESSO Then Error 16613
    
    'adiciona uma linha em Branco na combo Vinculo
     Vinculo.AddItem ""
        
    Set colCodigoNome = New AdmColCodigoNome

    'Leitura dos códigos e descrições das Regiões de Venda
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then Error 16614

    'Preenche a ComboBox Região com código e descrição das Regiões de Venda
    For Each objCodigoNome In colCodigoNome
        Regiao.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoNome.iCodigo
    Next

    Set colCodigo = New Collection

    'Lê cada código da tabela Estados
    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then Error 16615
    
    Call CF2(Me, "VND_Form_Load")
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16612, 16613, 16614, 16615, 16616, 25205

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158409)

    End Select

    iAlterado = 0
    
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
 
    Set objEventoTipoVendedor = Nothing
    Set objEventoBanco = Nothing
    Set objEventoSuperior = Nothing

    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub LabelBanco_Click()

Dim colSelecao As New Collection
Dim objBanco As New ClassBanco
    
    'Se o Banco estiver preenchido passa para o objBanco
    If Len(Trim(Banco.Text)) <> 0 Then objBanco.iCodBanco = CInt(Banco.Text)
    
    'Chama a tela que lista todos os Bancos
    Call Chama_Tela("BancoLista", colSelecao, objBanco, objEventoBanco)

End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then
    
        If Not IniciaLetra(NomeReduzido.Text) Then Error 25025
        
    End If

    'Preenche os labels vendedor existentes nos frames
    For iIndice = 0 To VendedorLabel.UBound
        VendedorLabel(iIndice).Caption = Trim(NomeReduzido.Text)
    Next
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True


    Select Case Err

        Case 25025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158410)

    End Select

End Sub

Private Sub objEventoBanco_evSelecao(obj1 As Object)
'retorno do browse de bancos

Dim lErro As Long
Dim objBanco As ClassBanco

On Error GoTo Erro_objEventoBanco_evSelecao

    Set objBanco = obj1

    If Not (objBanco Is Nothing) Then
        Banco.Text = CStr(objBanco.iCodBanco)
    End If

    Me.Show

    Exit Sub

Erro_objEventoBanco_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158411)

    End Select

    Exit Sub

End Sub

Public Sub PercComissao_Validate(Cancel As Boolean)

Dim lErro As String
Dim sPercComissao As String

On Error GoTo Erro_PercComissao_Validate

    'Verifica se foi preenchido a PercComissao
    If Len(Trim(PercComissao.Text)) = 0 Then Exit Sub

    'testa para ver se é uma porcentagem valida
    lErro = Porcentagem_Critica(PercComissao.Text)
    If lErro <> SUCESSO Then Error 33958

    sPercComissao = PercComissao.Text

    PercComissao.Text = Format(sPercComissao, "Fixed")

    Exit Sub

Erro_PercComissao_Validate:

    Cancel = True


    Select Case Err

        Case 33958

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158412)

    End Select

    Exit Sub

End Sub

Public Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'Verifica se foi preenchida a ComboBox Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
    If Regiao.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 16642

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tentativa de leitura da Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 16643
        
        'Não encontrou Regiao de Venda no BD
        If lErro = 16137 Then Error 16644

        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 16645

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 16642, 16643

        Case 16644  'Não encontrou Regiao de Venda no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

            If vbMsgRes = vbYes Then
                'Chama a tela RegiaoVenda
                Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
            End If

        Case 16645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", Err, Regiao.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158413)

    End Select

    Exit Sub


End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoVendedor As New ClassTipoVendedor
Dim iCodigo As Integer

On Error GoTo Erro_Tipo_Validate

    If iTipoAlterado = 0 Then Exit Sub
    
    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Tipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 16638

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoVendedor.iCodigo = iCodigo

        'Tentativa de leitura do Tipo Vendedor com esse código no BD
        lErro = CF("TipoVendedor_Le", objTipoVendedor)
        If lErro <> SUCESSO And lErro <> 16216 Then Error 16639

        If lErro = 16216 Then Error 16640 'Não encontrou Tipo Vendedor no BD

        'Encontrou Tipo Vendedor no BD, coloca no Text da Combo
        Tipo.Text = CStr(objTipoVendedor.iCodigo) & SEPARADOR & objTipoVendedor.sDescricao
        
        Call Trata_Tipo_Vendedor(objTipoVendedor)
        
    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 16641

    iTipoAlterado = 0
    
    Exit Sub

Erro_Tipo_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 16638, 16639

        Case 16640  'Não encontrou Tipo Vendedor no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPO_VENDEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela TipoVendedor
                Call Chama_Tela("TipoVendedor", objTipoVendedor)
            End If

        Case 16641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158414)

    End Select

    Exit Sub

End Sub

Public Sub TipoLabel_Click()
'chama browse de tipo de vendedor

Dim colSelecao As New Collection
Dim objTipoVendedor As New ClassTipoVendedor
    
    'Se o tipo de vendedor estiver preenchido --> passa seu código para o objTipoVendedor
    If Tipo.ListIndex <> -1 Then objTipoVendedor.iCodigo = Codigo_Extrai(Tipo.Text)
    
    'Chama a tela de Tipo de Vendedor
    Call Chama_Tela("TipoVendedorLista", colSelecao, objTipoVendedor, objEventoTipoVendedor)

End Sub

Private Sub objEventoTipoVendedor_evSelecao(obj1 As Object)
'retorno do browse de tipo de vendedor

Dim lErro As Long
Dim objTipoVendedor As ClassTipoVendedor
Dim bCancel As Boolean

On Error GoTo Erro_objEventoTipoVendedor_evSelecao

    Set objTipoVendedor = obj1

    If objTipoVendedor Is Nothing Then
        Tipo.Text = ""
    Else
        Tipo.Text = CStr(objTipoVendedor.iCodigo)
        Call Tipo_Validate(bCancel)
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoTipoVendedor_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158415)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Vendedores"

    'Le os dados da Tela Vendedores
    lErro = Le_Dados_Vendedor(objVendedor)
    If lErro <> SUCESSO Then Error 16617

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objVendedor.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objVendedor.sNome, STRING_VENDEDOR_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objVendedor.sNomeReduzido, STRING_VENDEDOR_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Endereco", objVendedor.lEndereco, 0, "Endereco"
    colCampoValor.Add "Tipo", objVendedor.iTipo, 0, "Tipo"
    colCampoValor.Add "Matricula", objVendedor.sMatricula, STRING_VENDEDOR_MATRICULA, "Matricula"
    colCampoValor.Add "CodRegiao", objVendedor.iCodRegiao, 0, "CodRegiao"
    colCampoValor.Add "PercComissao", objVendedor.dPercComissao, 0, "PercComissao"
    colCampoValor.Add "PercComissaoBaixa", objVendedor.dPercComissaoBaixa, 0, "PercComissaoBaixa"
    colCampoValor.Add "PercComissaoEmissao", objVendedor.dPercComissaoEmissao, 0, "PercComissaoEmissao"
    colCampoValor.Add "ComissaoSobreTotal", objVendedor.iComissaoSobreTotal, 0, "ComissaoSobreTotal"
    colCampoValor.Add "ComissaoFrete", objVendedor.iComissaoFrete, 0, "ComissaoFrete"
    colCampoValor.Add "ComissaoSeguro", objVendedor.iComissaoSeguro, 0, "ComissaoSeguro"
    colCampoValor.Add "ComissaoICM", objVendedor.iComissaoICM, 0, "ComissaoICM"
    colCampoValor.Add "ComissaoIPI", objVendedor.iComissaoIPI, 0, "ComissaoIPI"
    colCampoValor.Add "Banco", objVendedor.iBanco, 0, "Banco"
    colCampoValor.Add "Agencia", objVendedor.sAgencia, STRING_VENDEDOR_AGENCIA, "Agencia"
    colCampoValor.Add "ContaCorrente", objVendedor.sContaCorrente, STRING_VENDEDOR_CONTA_CORRENTE, "ContaCorrente"
    colCampoValor.Add "Ativo", objVendedor.iAtivo, 0, "ATivo"
    colCampoValor.Add "CGC", objVendedor.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objVendedor.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "RazaoSocial", objVendedor.sRazaoSocial, STRING_RAZAO_SOCIAL, "RazaoSocial"
    
    'incluido por Mauricio MAciel em 11/04/03
    colCampoValor.Add "Vinculo", objVendedor.iVinculo, 0, "Vinculo"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 16617

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158416)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Tela_Preenche

    objVendedor.iCodigo = colCampoValor.Item("Codigo").vValor

    If objVendedor.iCodigo <> 0 Then

        lErro = CF("Vendedor_Le_Estendida", objVendedor)
        If lErro <> SUCESSO And lErro <> 52705 Then gError 16619
        
        If lErro = SUCESSO Then
        
            lErro = Preenche_Tela_Vendedor(objVendedor)
            If lErro <> SUCESSO Then Error 16619
        
        End If
            
        VendedoresList.ListIndex = -1
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 16619, 52784 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158417)

    End Select

    Exit Sub
        
End Sub

Function Preenche_Tela_Vendedor(objVendedor As ClassVendedor) As Long
'Mostra os dados do Vendedor especificada em objVendedor

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim objTipoVendedor As New ClassTipoVendedor
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_Tela_Vendedor

    'PREENCHE LABELS DE VENDEDOR
'     VendedorLabel(0).Caption = objVendedor.sNomeReduzido
'     VendedorLabel(1).Caption = objVendedor.sNomeReduzido
    For iIndice = 0 To VendedorLabel.UBound
        VendedorLabel(iIndice).Caption = objVendedor.sNomeReduzido
    Next

    'MOSTRA IDENTIFICAÇÃO
    Codigo.Text = CStr(objVendedor.iCodigo)
    Nome.Text = objVendedor.sNome
    NomeReduzido.Text = objVendedor.sNomeReduzido
    Matricula.Text = objVendedor.sMatricula
    
    CGC.Text = objVendedor.sCgc
    Call CGC_Validate(bSGECancelDummy)
    CGC0.Format = CGC.Format
    CGC0.Text = CGC.Text
    
    RG.Text = objVendedor.sRG
    
    InscricaoEstadual.Text = objVendedor.sInscricaoEstadual
    RazaoSocial.Text = objVendedor.sRazaoSocial
     
    If objVendedor.iTipo = 0 Then
        
        Tipo.Text = ""
    
    Else
        
        Tipo.Text = CStr(objVendedor.iTipo)
        Call Tipo_Validate(bCancel)
            
    End If
    
    'Coloca código da Regiao em .Text e chama Validate
    If objVendedor.iCodRegiao <> 0 Then
        Regiao.Text = CStr(objVendedor.iCodRegiao)
        Call Regiao_Validate(bCancel)
    Else
        Regiao.Text = ""
    End If
        
    'MOSTRA COMISSÃO
    If objVendedor.dSaldoComissao <> 0 Then
        SaldoComissao.Caption = CStr(objVendedor.dSaldoComissao)
    Else
        SaldoComissao.Caption = ""
    End If
    
    If objVendedor.dPercComissao = 0 Then
        PercComissao.Text = ""
    Else
        PercComissao.Text = CStr(objVendedor.dPercComissao * 100)
    End If
    
    PercComissaoEmissao.Text = CStr(objVendedor.dPercComissaoEmissao * 100)
    PercComissaoBaixa.Caption = Format(objVendedor.dPercComissaoBaixa * 100, "#0.#0\%")
    
    ComissaoSobreTotal.Value = objVendedor.iComissaoSobreTotal
    ComissaoICM.Value = objVendedor.iComissaoICM
    ComissaoIPI.Value = objVendedor.iComissaoIPI
    ComissaoFrete.Value = objVendedor.iComissaoFrete
    ComissaoSeguro.Value = objVendedor.iComissaoSeguro
    
    If objVendedor.iBanco = 0 Then
        Banco.Text = ""
    Else
        Banco.Text = CStr(objVendedor.iBanco)
    End If
    
    Agencia.Text = objVendedor.sAgencia
    Ativo.Value = objVendedor.iAtivo
        
    ContaCorrente.Text = objVendedor.sContaCorrente
    
    lErro = Traz_Endereco_Tela(objVendedor)
    If lErro <> SUCESSO Then gError 12821
    
    'MOSTRA ESTATÍSTICA
    
    SaldoComissao.Caption = Format(objVendedor.dSaldoComissao, "Standard")
    
    If objVendedor.dtDataUltimaVenda <> DATA_NULA Then
        DataUltVenda.Caption = Format(objVendedor.dtDataUltimaVenda, "dd/mm/yyyy")
    Else
        DataUltVenda.Caption = "  /  /    "
    End If
    
    'Preenche a combo Vinculo com codigo de Vinculo passado em objVendedor - incluido por Mauricio Maciel em 11/04/03
    For iIndice = 0 To Vinculo.ListCount - 1
    
        'verifica se o item é o que esta sendo procurado
        If Vinculo.ItemData(iIndice) = objVendedor.iVinculo Then
            
            Vinculo.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
    
    'Coloca Tipo no Text
    If objVendedor.iCargo <> 0 Then
        Cargo.Text = CStr(objVendedor.iCargo)
    
        'Tenta selecionar
        lErro = Combo_Seleciona(Cargo, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195878
    Else
        Cargo.ListIndex = -1
    End If
    
    If objVendedor.iSuperior = 0 Then
        Superior.Text = ""
    Else
        Superior.Text = CStr(objVendedor.iSuperior)
        Call Superior_Validate(bCancel)
    End If
    
    If objVendedor.sCodUsuario <> "" Then
        For iIndice = 0 To CodUsuario.ListCount - 1
            If objVendedor.sCodUsuario = CodUsuario.List(iIndice) Then
                CodUsuario.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        CodUsuario.ListIndex = -1
    End If
    
    lErro = CF2(Me, "VND_Traz_Vendedor_Tela", objVendedor)
    If lErro <> SUCESSO Then gError 195876
    
    iAlterado = 0
    
    Preenche_Tela_Vendedor = SUCESSO
    
    Exit Function
    
Erro_Preenche_Tela_Vendedor:
    
    Preenche_Tela_Vendedor = gErr
    
    Select Case gErr
    
        Case 12821, 195876, 195878

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158418)

    End Select

    Exit Function
    
End Function

Function Traz_Endereco_Tela(objVendedor As ClassVendedor) As Long
'Traz os dados para tela

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim colEnderecos As New Collection

On Error GoTo Erro_Traz_Endereco_Tela
   
    objEndereco.lCodigo = objVendedor.lEndereco
    
    'Le endereço no BD
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then Error 16620
  
    colEnderecos.Add objEndereco

    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 16620

    Traz_Endereco_Tela = SUCESSO

Exit Function

Erro_Traz_Endereco_Tela:

    Traz_Endereco_Tela = Err

    Select Case Err
    
        Case 16620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158419)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Vendedor(objVendedor As ClassVendedor) As Long
'Lê os dados que estão na tela Vendedores e coloca em objVendedor

Dim iPosicao As Integer
   
    objVendedor.iCodigo = StrParaInt(Codigo.Text)
    objVendedor.sNome = Nome.Text
    objVendedor.sNomeReduzido = NomeReduzido.Text
    objVendedor.sCodUsuario = CodUsuario.Text

    If Len(Trim(Tipo.Text)) = 0 Then
        objVendedor.iTipo = 0
    Else
        objVendedor.iTipo = Codigo_Extrai(Tipo.Text)
    End If
    
    If Len(Trim(Matricula.Text)) = 0 Then
        objVendedor.sMatricula = ""
    Else
        objVendedor.sMatricula = Matricula.Text
    End If
    
    objVendedor.sInscricaoEstadual = Trim(InscricaoEstadual.Text)
    objVendedor.sRazaoSocial = Trim(RazaoSocial.Text)

    If Len(Trim(Regiao.Text)) = 0 Then
        objVendedor.iCodRegiao = 0
    Else
        objVendedor.iCodRegiao = Codigo_Extrai(Regiao.Text)
    End If
    
    If Len(Trim(PercComissaoEmissao.Text)) = 0 Then
        objVendedor.dPercComissaoEmissao = 0
    Else
        objVendedor.dPercComissaoEmissao = CDbl(PercComissaoEmissao.Text) / 100
    End If
    
    If Trim(PercComissaoBaixa.Caption) = "" Then
        objVendedor.dPercComissaoBaixa = 0
    Else
        objVendedor.dPercComissaoBaixa = PercentParaDbl(PercComissaoBaixa.Caption)
    End If
    
    If Len(Trim(PercComissao.Text)) = 0 Then
        objVendedor.dPercComissao = 0
    Else
        objVendedor.dPercComissao = CDbl(PercComissao.Text) / 100
    End If
    
    With objVendedor

        .iComissaoSobreTotal = ComissaoSobreTotal.Value
        .iComissaoFrete = ComissaoFrete.Value
        .iComissaoICM = ComissaoICM.Value
        .iComissaoIPI = ComissaoIPI.Value
        .iComissaoSeguro = ComissaoSeguro.Value

    End With
        
    objVendedor.sContaCorrente = ContaCorrente.Text
    objVendedor.sAgencia = Agencia.Text
    objVendedor.iAtivo = Ativo.Value
    objVendedor.iBanco = StrParaInt(Banco.Text)
    
    'inclusão feita por Mauricio Maciel em 11/04/03
    If Vinculo.ListIndex <> -1 Then
        objVendedor.iVinculo = Vinculo.ItemData(Vinculo.ListIndex)
    Else
        objVendedor.iVinculo = 0
    End If
    
    If objVendedor.iVinculo = 1 Then
        objVendedor.sCgc = Trim(CGC0.Text)
    Else
        objVendedor.sCgc = Trim(CGC.Text)
    End If
    objVendedor.sRG = Trim(RG.Text)
    
    Le_Dados_Vendedor = SUCESSO
    
End Function

Function Trata_Parametros(Optional objVendedor As ClassVendedor) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se existir um Vendedor passado como parametro, exibir seus dados
    If Not (objVendedor Is Nothing) Then

        lErro = CF("Vendedor_Le_Estendida", objVendedor)
        If lErro <> SUCESSO And lErro <> 52705 Then gError 16621
        
        'Se nao encontrou pelo código => tenta pelo nome reduzido
        If lErro <> SUCESSO Then
        
            'Alteracao Daniel
            lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
            If lErro <> SUCESSO And lErro <> 25008 Then gError 97900
            'Fim da Alteracao Daniel
            
        End If
        
        If lErro = SUCESSO Then

            'Mostra dados na tela
            lErro = Preenche_Tela_Vendedor(objVendedor)
            If lErro <> SUCESSO Then gError 16622

        Else
            
            'Limpa a tela
            Call Limpa_Tela_Vendedor
            
            If objVendedor.iCodigo <> 0 Then
            
                
                'Exibe apenas o codigo
                Codigo.Text = CStr(objVendedor.iCodigo)
            
            End If
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 16621, 16622, 58240, 97900 'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158420)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Limpa_Tela_Vendedor()
'Limpa todos os campos de input da tela Vendedores

Dim lErro As Long
Dim iIndice As Integer

    Call Limpa_Tela(Me)
    
    Codigo.Text = ""
    Ativo.Value = MARCADO
    
    'Desmarca ListBox VendedoresList
    VendedoresList.ListIndex = -1
    
    Cargo.ListIndex = -1
    CodUsuario.ListIndex = -1
    
    'Desmarca Labels, ComboBoxes e checkBoxes
    PercComissaoBaixa.Caption = ""
    SaldoComissao.Caption = ""
'    VendedorLabel(0).Caption = ""
'    VendedorLabel(1).Caption = ""
    For iIndice = 0 To VendedorLabel.UBound
        VendedorLabel(iIndice).Caption = ""
    Next
    
    Tipo.ListIndex = -1
    iTipoAlterado = 0
    
    Regiao.ListIndex = -1
    
    'Incluido por Mauricio Maciel em 11/04/03
    Vinculo.ListIndex = -1
        
    'Limpa os dados do Frame de Estatísticas
    SaldoComissao.Caption = ""
    
    DataUltVenda.Caption = "  /  /    "
    
    ComissaoFrete.Value = vbUnchecked
    ComissaoICM.Value = vbUnchecked
    ComissaoIPI.Value = vbUnchecked
    ComissaoSeguro.Value = vbUnchecked
    
    Call gobjTabEnd.Limpa_Tela
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call CF2(Me, "VND_Limpa_Tela_Vendedor")
    
    iAlterado = 0
    
End Function

Public Sub Matricula_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

'Incluido por Mauricio MAciel em 11/04/03
Public Sub Vinculo_Click()
Dim iIndice As Integer

    iAlterado = REGISTRO_ALTERADO
    
    For iIndice = 0 To 2
        FrameInfoVinculo(iIndice).Visible = (iIndice = Vinculo.ListIndex)
    Next
    
End Sub

Public Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame selecionada visível, esconde a frame atual
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_VENDEDORES_ID
                
            Case TAB_Comissao
                Parent.HelpContextID = IDH_VENDEDORES_COMISSAO
                        
            Case TAB_Endereco
                Parent.HelpContextID = IDH_VENDEDORES_ENDERECO
        
        End Select
        
        
    End If

End Sub

Public Sub PercComissao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Regiao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Regiao_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Telefone1_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Telefone2_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO
    iTipoAlterado = 1
    
End Sub

Public Sub Tipo_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_Tipo_Click

    If Tipo.ListIndex <> -1 Then
    
        objTipoVendedor.iCodigo = Tipo.ItemData(Tipo.ListIndex)
        
        lErro = CF("TipoVendedor_Le", objTipoVendedor)
        If lErro <> SUCESSO And lErro <> 16216 Then Error 16637
        
        If lErro = SUCESSO Then 'Tipo existe
        
            Call Trata_Tipo_Vendedor(objTipoVendedor)
            
        Else 'Tipo não existe
             vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPO_VENDEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de TipoVendedor
                Call Chama_Tela("TipoVendedor", objTipoVendedor)
            Else
                Tipo.SetFocus
            End If
            
        End If
        
    End If
    
    iAlterado = REGISTRO_ALTERADO
        
    Exit Sub
    
Erro_Tipo_Click:

    Select Case Err
    
        Case 16637
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158422)
        
    End Select
    
    Exit Sub
        
End Sub

Public Sub VendedoresList_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedoresList_DblClick

    'Extrai código do Vendedor
    objVendedor.iCodigo = VendedoresList.ItemData(VendedoresList.ListIndex)

    'Lê Vendedor a partir do código
    lErro = CF("Vendedor_Le_Estendida", objVendedor)
    If lErro <> SUCESSO And lErro <> 12582 Then Error 16634
    
    If lErro <> SUCESSO Then Error 16636 'Vendedor não cadastrado

    'Mostra dados na tela
    lErro = Preenche_Tela_Vendedor(objVendedor)
    If lErro <> SUCESSO Then Error 16635

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_VendedoresList_DblClick:

    Select Case Err

        Case 16634, 16635

        Case 16636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158423)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub PercComissaoEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercComissaoEmissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercComissaoEmissao As Double
Dim dPercComissaoBaixa As Double

On Error GoTo Erro_PercComissaoEmissao_Validate

    If Len(Trim(PercComissaoEmissao.Text)) <> 0 Then

        lErro = Porcentagem_Critica(PercComissaoEmissao.Text)
        If lErro <> SUCESSO Then Error 16648

        PercComissaoEmissao.Text = Format(PercComissaoEmissao.Text, "Fixed")

        dPercComissaoEmissao = CDbl(PercComissaoEmissao.Text)

        dPercComissaoBaixa = 100 - dPercComissaoEmissao

        PercComissaoBaixa.Caption = Format(dPercComissaoBaixa, "#0.#0\%")

    End If

    Exit Sub

Erro_PercComissaoEmissao_Validate:

    Cancel = True

    Select Case Err

        Case 16648

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158424)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_VENDEDORES_ID
    Set Form_Load_Ocx = Me
    Caption = "Vendedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Vendedores"
    
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
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Banco Then
            Call LabelBanco_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call TipoLabel_Click
        ElseIf Me.ActiveControl Is Superior Then
            Call LabelSuperior_Click
        End If
    End If

End Sub

Private Sub Trata_Tipo_Vendedor(objTipoVendedor As ClassTipoVendedor)

    'Preenche dados da tela correspondente ao tipo selecionado
    PercComissao.Text = CStr(objTipoVendedor.dPercComissao * 100)
    PercComissaoEmissao.Text = CStr(objTipoVendedor.dPercComissaoEmissao * 100)
    PercComissaoBaixa.Caption = Format(objTipoVendedor.dPercComissaoBaixa * 100, "#0.#0\%")
    ComissaoSobreTotal.Value = objTipoVendedor.iComissaoSobreTotal
    ComissaoICM.Value = objTipoVendedor.iComissaoICM
    ComissaoIPI.Value = objTipoVendedor.iComissaoIPI
    ComissaoFrete.Value = objTipoVendedor.iComissaoFrete
    ComissaoSeguro.Value = objTipoVendedor.iComissaoSeguro

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

Public Sub RazaoSocial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CGC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub InscricaoEstadual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Function VND_Form_Load(ByVal objCT As Object) As Long
    VND_Form_Load = SUCESSO
End Function

Public Function VND_Grava_Registro(ByVal objCT As Object, ByVal objVendedor As ClassVendedor) As Long
    VND_Grava_Registro = SUCESSO
End Function

Public Function VND_Traz_Vendedor_Tela(ByVal objCT As Object, ByVal objVendedor As ClassVendedor) As Long
    VND_Traz_Vendedor_Tela = SUCESSO
End Function

Public Function VND_Limpa_Tela_Vendedor(ByVal objCT As Object) As Long
    VND_Limpa_Tela_Vendedor = SUCESSO
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = CF2(Me, "VND_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 195899
                
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 195900

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 195899

        Case 105900
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155604)

    End Select

    Exit Function

End Function

Public Function VND_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    VND_Saida_Celula = SUCESSO
End Function

Public Sub Superior_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Superior_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Superior_Validate

    If Len(Trim(Superior.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Superior, objVendedor)
        If lErro <> SUCESSO Then gError 195884
        
        If objVendedor.iAtivo = DESMARCADO Then gError 195885
        
    End If

    Exit Sub

Erro_Superior_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195884
        
        Case 195885
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195886)
    
    End Select

End Sub

Public Sub LabelSuperior_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Superior.Text)) > 0 Then objVendedor.sNomeReduzido = Superior.Text
    
    Set objEventoSuperior = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoSuperior)

End Sub

Private Sub objEventoSuperior_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Superior.Text = objVendedor.sNomeReduzido

    iAlterado = 0

    Me.Show

    Exit Sub

End Sub

Public Sub Cargo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cargo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cargo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cargo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_CARGO_VENDEDOR, Cargo, "AVISO_CRIAR_CARGO_VENDEDOR")
    If lErro <> SUCESSO Then gError 195867
    
    Exit Sub

Erro_Cargo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 195867
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195868)

    End Select

End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long, colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    CodUsuario.AddItem ""
    
    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then Error 48100

    For Each objUsuarios In colUsuarios
        CodUsuario.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = Err

    Select Case Err

        Case 48100

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142643)

    End Select

    Exit Function

End Function

Public Sub CodUsuario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CodUsuario_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
    
Public Sub RG_GotFocus()
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)
End Sub

Public Sub CGC_GotFocus()
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)
End Sub

Public Sub CGC0_GotFocus()
    Call MaskEdBox_TrataGotFocus(CGC0, iAlterado)
End Sub

Public Sub CGC0_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RG_Change()
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

Public Sub CGC0_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC0_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC0.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC0.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC0.Text)
            If lErro <> SUCESSO Then Error 12316
            
            'Formata e coloca na Tela
            CGC0.Format = "000\.000\.000-00; ; ; "
            CGC0.Text = CGC0.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC0.Text)
            If lErro <> SUCESSO Then Error 12317
            
            'Formata e Coloca na Tela
            CGC0.Format = "00\.000\.000\/0000-00; ; ; "
            CGC0.Text = CGC0.Text

        Case Else
                
            Error 12318

    End Select

    Exit Sub

Erro_CGC0_Validate:

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

