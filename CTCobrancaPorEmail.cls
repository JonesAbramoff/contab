VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCobrancaPorEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object 'Inserido por wagner

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoBorderoDe As AdmEvento
Attribute objEventoBorderoDe.VB_VarHelpID = -1

Public iAlterado As Integer
Public iAlteradoFiltro As Integer

Public iFrameAtual As Integer

Public iLinhaAnt As Integer

Public giTipoTela As Integer

'Grid de Itens
Public objGridItens As AdmGrid
Public iGrid_Selecionado_Col As Integer
Public iGrid_Numero_Col As Integer
Public iGrid_Parcela_Col As Integer
Public iGrid_Tipo_Col As Integer
Public iGrid_Cliente_Col As Integer
Public iGrid_Filial_Col As Integer
Public iGrid_DataVencimento_Col As Integer
Public iGrid_Valor_Col As Integer
Public iGrid_Saldo_Col As Integer
Public iGrid_Atraso_Col As Integer
Public iGrid_Carta_Col As Integer
Public iGrid_CC_Col As Integer
Public iGrid_Assunto_Col As Integer
Public iGrid_Email_Col As Integer
Public iGrid_Mensagem_Col As Integer
Public iGrid_Anexo_Col As Integer

Public gcolParcelas As Collection
Public gcolTitulos As Collection
Public gcolClientes As Collection
Public gcolFiliaisCli As Collection
Public gcolModeloLinha As Collection
Public gcolEndereco As Collection

Public giLinhaAtual As Integer

Public gcolParcelasEnv As Collection
Public gcolTitulosEnv As Collection
Public gcolClientesEnv As Collection
Public gcolFiliaisCliEnv As Collection
Public gcolEnderecoEnv As Collection

Const TAB_SELECAO = 1
Const TAB_DADOS = 2

'Mnemonicos
Const DATA_VENCIMENTO = "Data_Vencimento"
Const DATA_VENCIMENTO_REAL = "Data_Vencimento_Real"
Const NUMERO_PARCELA = "Numero_Parcela"
Const NUMERO_TITULO = "Numero_Titulo"
Const NUMERO_TITULO_TRV = "Numero_Titulo_TRV"
Const NUMERO_TITULO_TRP = "Numero_Titulo_TRP"
Const RAZAO_CLIENTE = "Razao_Cliente"
Const CNPJ_CLIENTE = "CNPJ_Cliente"
Const SALDO_PARCELA = "Saldo_Parcela"
Const VALOR_PARCELA = "Valor_Parcela"
Const DIA_ATUAL = "Dia_Atual"
Const MES_NOME = "Mes_Nome"
Const MES_ATUAL = "Mes_Atual"
Const ANO_ATUAL = "Ano_Atual"
Const DATA_ATUAL = "Data_Atual"
Const DATA_BAIXA = "Data_Baixa"
Const NOME_EMP = "Nome_Emp"
Const CNPJ_EMP = "CNPJ_Emp"
Const ENDERECO_EMP = "Endereco_Emp"
Const BAIRRO_EMP = "Bairro_Emp"
Const CIDADE_EMP = "Cidade_Emp"
Const UF_EMP = "UF_Emp"
Const CEP_EMP = "Cep_Emp"
Const ARQ_XML = "NomeArq_XML"
Const DIR_XML = "DirArq_XML"
Const NFE_CHAVE = "NFe_Chave"
Const CAMINHO_XML = "CaminhoComplArq_XML"
Const VALOR_TOTAL = "Valor_Total"
Const NUMERO_NFE = "Numero_NFE"
Const DATA_EMISSAO = "Data_Emissao"
Const LISTA_NFSPAG = "Lista_NFsPag"
Const BOLETO_CORPORATOR = "Boleto_Corporator"
Const EMAIL_TRANSP_SAIDA = "Email_Transp_Saida"
Const EMAIL_VENDEDOR = "Email_Vendedor"
Const IM_EMP = "IM_Emp"
Const IM_EMP2 = "IM_Emp2"
Const NFSE_LINK_CONSULTA = "NFSe_Link_Consulta"
Const NFSE_LINK_SITE = "NFSe_Link_Site"
Const NFSE_LINK_VERIFICACAO = "NFSe_LinkVerificacao"
Const NFE_CHAVE2 = "NFe_Chave2"
Const NFSE_LINKS_PADRAO = "NFSe_Links_Padrao"

'--- inicio dos properties get dos controles da tela
Public Property Get optIncluiDanfe() As Object
     Set optIncluiDanfe = objUserControl.Controls("optIncluiDanfe")
End Property

Public Property Get cCliente() As Object
     Set cCliente = objUserControl.Controls("cCliente")
End Property

Public Property Get cTransportadora() As Object
     Set cTransportadora = objUserControl.Controls("cTransportadora")
End Property

Public Property Get cFixo() As Object
     Set cFixo = objUserControl.Controls("cFixo")
End Property

Public Property Get emailFixo() As Object
     Set emailFixo = objUserControl.Controls("emailFixo")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get AnexoGrid() As Object
     Set AnexoGrid = objUserControl.Controls("AnexoGrid")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get Anexo() As Object
     Set Anexo = objUserControl.Controls("Anexo")
End Property

Public Property Get Email() As Object
     Set Email = objUserControl.Controls("Email")
End Property

Public Property Get Cc() As Object
     Set Cc = objUserControl.Controls("Cc")
End Property

Public Property Get Assunto() As Object
     Set Assunto = objUserControl.Controls("Assunto")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get LabelCc() As Object
     Set LabelCc = objUserControl.Controls("LabelCc")
End Property

Public Property Get MensagemGrid() As Object
     Set MensagemGrid = objUserControl.Controls("MensagemGrid")
End Property

Public Property Get AssuntoGrid() As Object
     Set AssuntoGrid = objUserControl.Controls("AssuntoGrid")
End Property

Public Property Get CCGrid() As Object
     Set CCGrid = objUserControl.Controls("CCGrid")
End Property

Public Property Get EmailGrid() As Object
     Set EmailGrid = objUserControl.Controls("EmailGrid")
End Property

Public Property Get Carta() As Object
     Set Carta = objUserControl.Controls("Carta")
End Property

Public Property Get Atraso() As Object
     Set Atraso = objUserControl.Controls("Atraso")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get Selecionado() As Object
     Set Selecionado = objUserControl.Controls("Selecionado")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get BotaoMarcarTodos() As Object
     Set BotaoMarcarTodos = objUserControl.Controls("BotaoMarcarTodos")
End Property

Public Property Get BotaoDesmarcarTodos() As Object
     Set BotaoDesmarcarTodos = objUserControl.Controls("BotaoDesmarcarTodos")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get TituloInic() As Object
     Set TituloInic = objUserControl.Controls("TituloInic")
End Property

Public Property Get TituloFim() As Object
     Set TituloFim = objUserControl.Controls("TituloFim")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get SaldoDe() As Object
     Set SaldoDe = objUserControl.Controls("SaldoDe")
End Property

Public Property Get SaldoAte() As Object
     Set SaldoAte = objUserControl.Controls("SaldoAte")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get TipoDocSeleciona() As Object
     Set TipoDocSeleciona = objUserControl.Controls("TipoDocSeleciona")
End Property

Public Property Get TipoDocTodos() As Object
     Set TipoDocTodos = objUserControl.Controls("TipoDocTodos")
End Property

Public Property Get TipoDocApenas() As Object
     Set TipoDocApenas = objUserControl.Controls("TipoDocApenas")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get AtrasoDe() As Object
     Set AtrasoDe = objUserControl.Controls("AtrasoDe")
End Property

Public Property Get AtrasoAte() As Object
     Set AtrasoAte = objUserControl.Controls("AtrasoAte")
End Property

Public Property Get DataBaixaDe() As Object
     Set DataBaixaDe = objUserControl.Controls("DataBaixaDe")
End Property

Public Property Get DataBaixaAte() As Object
     Set DataBaixaAte = objUserControl.Controls("DataBaixaAte")
End Property

Public Property Get FrameCliente() As Object
     Set FrameCliente = objUserControl.Controls("FrameCliente")
End Property

Public Property Get ClienteInicial() As Object
     Set ClienteInicial = objUserControl.Controls("ClienteInicial")
End Property

Public Property Get ClienteFinal() As Object
     Set ClienteFinal = objUserControl.Controls("ClienteFinal")
End Property

Public Property Get LabelClienteDe() As Object
     Set LabelClienteDe = objUserControl.Controls("LabelClienteDe")
End Property

Public Property Get LabelClienteAte() As Object
     Set LabelClienteAte = objUserControl.Controls("LabelClienteAte")
End Property

Public Property Get BotaoEmail() As Object
     Set BotaoEmail = objUserControl.Controls("BotaoEmail")
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

Public Property Get Modelo() As Object
     Set Modelo = objUserControl.Controls("Modelo")
End Property

Public Property Get EmailValido() As Object
     Set EmailValido = objUserControl.Controls("EmailValido")
End Property

Public Property Get IgnoraJaEnviados() As Object
     Set IgnoraJaEnviados = objUserControl.Controls("IgnoraJaEnviados")
End Property

Public Property Get CategoriaClienteTodas() As Object
     Set CategoriaClienteTodas = objUserControl.Controls("CategoriaClienteTodas")
End Property

Public Property Get CategoriaCliente() As Object
     Set CategoriaCliente = objUserControl.Controls("CategoriaCliente")
End Property

Public Property Get CategoriaClienteDe() As Object
     Set CategoriaClienteDe = objUserControl.Controls("CategoriaClienteDe")
End Property

Public Property Get CategoriaClienteAte() As Object
     Set CategoriaClienteAte = objUserControl.Controls("CategoriaClienteAte")
End Property

Public Property Get comboCobrador() As Object
     Set comboCobrador = objUserControl.Controls("comboCobrador")
End Property

Public Property Get ComboCobradorBordero() As Object
     Set ComboCobradorBordero = objUserControl.Controls("ComboCobradorBordero")
End Property

Public Property Get BorderoCobrDe() As Object
     Set BorderoCobrDe = objUserControl.Controls("BorderoCobrDe")
End Property

'--- fim dos properties get dos controles da tela

Public Sub AtrasoAte_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub AtrasoDe_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185997)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        Caption = "Cobrança do envio de Fatura"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Then
        Caption = "Agradecimento por pagamento"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        Caption = "Aviso de pagamento de fatura"
    ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
        Caption = "Cobrança de títulos atrasados"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        Caption = "Aviso de títulos a vencer"
    Else
        Caption = "Nota Fiscal Eletrônica Federal"
    End If
    Call Form_Load

End Function

Public Function Name() As String

    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        Name = "CobrancaFaturaPorEmail"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Then
        Name = "AgradecimentoPorEmail"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        Name = "AvisoPagtoCPPorEmail"
    ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
        Name = "CobrancaPorEmail"
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        Name = "AvisoCobrPorEmail"
    Else
        Name = "NFePorEmail"
    End If

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

'Public Property Get Parent() As Object
'    Set Parent = UserControl.Parent
'End Property
'**** fim do trecho a ser copiado *****

Public Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim objCobrancaEmailSel As New ClassCobrancaPorEmailSel
Dim objEmailConfig As New ClassEmailConfig
Dim iFilialEmpresa As Integer

On Error GoTo Erro_TabStrip1_BeforeClick

    GL_objMDIForm.MousePointer = vbHourglass

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
        
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_SELECAO Then
    
        'Valida a seleção
        lErro = ValidaSelecao()
        If lErro <> SUCESSO Then gError 185998
        
        If iAlteradoFiltro = REGISTRO_ALTERADO Then
        
            objEmailConfig.sUsuario = gsUsuario
            
            lErro = CF("EmailConfig_Le", objEmailConfig)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202829
            
            If lErro <> SUCESSO Then gError 202830
        
            lErro = Move_TabSelecao_Memoria(objCobrancaEmailSel)
            If lErro <> SUCESSO Then gError 185999
                                       
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
                lErro = CF("CobrancaPorEmail_NFsPag_Le", objCobrancaEmailSel)
            ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
                lErro = CF("CobrancaPorEmail_ParcPagBaixados_Le", objCobrancaEmailSel)
            ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Then
                lErro = CF("CobrancaPorEmail_ParcelasBaixadas_Le", objCobrancaEmailSel)
            ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
                lErro = CF("CobrancaPorEmail_ParcelasEmAtraso_Le", objCobrancaEmailSel)
            ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
                iFilialEmpresa = giFilialEmpresa
                If giFilialEmpresa > 50 Then giFilialEmpresa = giFilialEmpresa - 50
                lErro = CF("CobrancaPorEmail_NFe_Le", objCobrancaEmailSel)
                giFilialEmpresa = iFilialEmpresa
            End If
            If lErro <> SUCESSO Then gError 187000
            
            If objCobrancaEmailSel.colParcelas.Count = 0 Then gError 187117
                            
            lErro = Preenche_GridItens(objCobrancaEmailSel)
            If lErro <> SUCESSO Then gError 187001
            
            iAlteradoFiltro = 0
            
        End If
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True
    
    If iFilialEmpresa <> 0 Then giFilialEmpresa = iFilialEmpresa

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 185998, 185999, 187000, 187001, 202829
        
        Case 187117
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_HISTCOBRANCA_SEM_PARCELAS", gErr)

        Case 202830
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCONFIG_NAO_CADASTRADO", gErr, objEmailConfig.sUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187002)

    End Select

    Exit Sub

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
                              
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187003)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCobranca As New Collection
Dim objCobranca As ClassCobrancaEmailPadrao
Dim iTipo As Integer

On Error GoTo Erro_Form_Load
        
    iFrameAtual = 1
    
    Set gcolParcelas = New Collection
    Set gcolTitulos = New Collection
    Set gcolClientes = New Collection
    Set gcolFiliaisCli = New Collection
    Set gcolModeloLinha = New Collection
    Set gcolEndereco = New Collection
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoBorderoDe = New AdmEvento
    
    'Grid Itens
    Set objGridItens = New AdmGrid

    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 187004
    
    If giTipoTela = TIPOTELA_EMAIL_NFE Then
        Call DateParaMasked(DataBaixaDe, gdtDataAtual)
    End If
    
    If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
        lErro = Carrega_TipoDocumento(TipoDocSeleciona)
        If lErro <> SUCESSO Then gError 187005
    End If
        
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
        
        Call Carrega_ComboCategoriaCliente(CategoriaCliente)
    
        '#####################################
        'Inserido por Wagner
        CategoriaClienteTodas.Value = vbChecked
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
        CategoriaClienteDe.ListIndex = -1
        CategoriaClienteAte.ListIndex = -1
        '#####################################
        
    End If
    
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Or giTipoTela = TIPOTELA_EMAIL_NFE Then
        
        If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Then
            iTipo = TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO
        ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
            iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA
        ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
            iTipo = TIPO_COBRANCAEMAILPADRAO_NFE
        Else
            iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP
        End If
    
        Carta.Clear
        
        'Le os modelos válidos para o atraso em questão
        lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, iTipo)
        If lErro <> SUCESSO Then gError 189429
            
        'Carrega a Combo com os Dados da Colecao
        For Each objCobranca In colCobranca
            Carta.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
            Carta.ItemData(Carta.NewIndex) = objCobranca.lCodigo
        Next
        
        'Carrega a Combo com os Dados da Colecao
        Modelo.Clear
        For Each objCobranca In colCobranca
            Modelo.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
            Modelo.ItemData(Modelo.NewIndex) = objCobranca.lCodigo
        Next
        
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
    
        Carta.Clear
    
        'Le os modelos válidos para o atraso em questão
        lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_AVISO)
        If lErro <> SUCESSO Then gError 189429
        
        'Carrega a Combo com os Dados da Colecao
        For Each objCobranca In colCobranca
            Carta.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
            Carta.ItemData(Carta.NewIndex) = objCobranca.lCodigo
        Next
                
        'Carrega a Combo com os Dados da Colecao
        Modelo.Clear
        For Each objCobranca In colCobranca
            Modelo.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
            Modelo.ItemData(Modelo.NewIndex) = objCobranca.lCodigo
        Next
        
        lErro = Carrega_Usuarios
        If lErro <> SUCESSO Then gError 187004
        
        lErro = Carrega_Combo_CobradorBordero
        If lErro <> SUCESSO Then gError 187004
        
    ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
        
        'Le os modelos válidos para o atraso em questão
        lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_COBRANCA)
        If lErro <> SUCESSO Then gError 189429
        
        'Carrega a Combo com os Dados da Colecao
        Modelo.Clear
        For Each objCobranca In colCobranca
            Modelo.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
            Modelo.ItemData(Modelo.NewIndex) = objCobranca.lCodigo
        Next
        
        lErro = Carrega_Usuarios
        If lErro <> SUCESSO Then gError 187004
        
        lErro = Carrega_Combo_CobradorBordero
        If lErro <> SUCESSO Then gError 187004
        
    End If
    
    iAlterado = 0
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 187004, 187005, 189429
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187006)

    End Select

    Exit Sub

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoBorderoDe = Nothing
    
    Set gcolParcelas = Nothing
    Set gcolTitulos = Nothing
    Set gcolClientes = Nothing
    Set gcolFiliaisCli = Nothing
    Set gcolModeloLinha = Nothing
    Set gcolEndereco = Nothing
    
    Set gcolParcelasEnv = Nothing
    Set gcolTitulosEnv = Nothing
    Set gcolClientesEnv = Nothing
    Set gcolFiliaisCliEnv = Nothing
    Set gcolEnderecoEnv = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187007)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal iTipoTela As Integer = TIPOTELA_EMAIL_COBRANCA, Optional ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giTipoTela = iTipoTela
    
    If Not (objCobrancaEmailSel Is Nothing) Then
    
        'Torna Frame atual invisível
        Frame1(TabStrip1.SelectedItem.Index).Visible = False
        iFrameAtual = TAB_DADOS
        'Torna Frame atual visível
        Frame1(iFrameAtual).Visible = True
        TabStrip1.Tabs.Item(iFrameAtual).Selected = True
        
        iAlteradoFiltro = 0
        
        lErro = Preenche_GridItens(objCobrancaEmailSel)
        If lErro <> SUCESSO Then gError 192212
    
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 192212

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187008)

    End Select
    
    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    colcolColecoes.Add gcolClientes
    colcolColecoes.Add gcolFiliaisCli
    colcolColecoes.Add gcolParcelas
    colcolColecoes.Add gcolTitulos
    colcolColecoes.Add gcolModeloLinha
    colcolColecoes.Add gcolEndereco
    
    Call Ordenacao_ClickGrid(objGridItens, Nothing, colcolColecoes)

End Sub

Public Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Public Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Public Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

    Call Recolhe_Dados_Email(iLinhaAnt)
    Call Mostra_Dados_Email(GridItens.Row)
    
    iLinhaAnt = GridItens.Row

End Sub

Public Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_Carta_Col
                
                    lErro = Saida_Celula_Carta(objGridInt)
                    If lErro <> SUCESSO Then gError 187009
                
                Case Else
                    'Não há tratamento específico na saída de célula,
                    'uma vez que os campos não serão editados
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 187010

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 187009

        Case 187010
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187011)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Carta(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lCarta As Long
Dim lCartaAnt As Long

On Error GoTo Erro_Saida_Celula_Carta

    Set objGridInt.objControle = Carta
    
    lCarta = LCodigo_Extrai(Carta.Text)
    lCartaAnt = LCodigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Carta_Col))
    
    lErro = Preenche_Dados_Carta(lCarta, lCartaAnt, GridItens.Row)
    If lErro <> SUCESSO Then gError 187012
    
    Call Mostra_Dados_Email(GridItens.Row)
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 187013

    Saida_Celula_Carta = SUCESSO

    Exit Function

Erro_Saida_Celula_Carta:

    Saida_Celula_Carta = gErr

    Select Case gErr

        Case 187012, 187013
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 187014)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add (" ")

    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        objGrid.colColuna.Add ("Número")
    ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
        objGrid.colColuna.Add ("Série")
        objGrid.colColuna.Add ("Número")
        objGrid.colColuna.Add ("Tipo")
    Else
        objGrid.colColuna.Add ("Título")
        objGrid.colColuna.Add ("Parcela")
        objGrid.colColuna.Add ("Tipo")
    End If
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    
    If giTipoTela = TIPOTELA_EMAIL_NFE Then
        objGrid.colColuna.Add ("Emissão")
    Else
        objGrid.colColuna.Add ("Vencimento")
    End If
    
    objGrid.colColuna.Add ("Valor")
    
    If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
        objGrid.colColuna.Add ("Saldo")
    End If
    
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
        objGrid.colColuna.Add ("Atraso(dias)")
    ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        objGrid.colColuna.Add ("Baixa")
    ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        objGrid.colColuna.Add ("Emissão")
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        objGrid.colColuna.Add ("Vence em(dias)")
    End If
    
    objGrid.colColuna.Add ("Modelo")
    objGrid.colColuna.Add ("Email")
    objGrid.colColuna.Add ("CC")
    objGrid.colColuna.Add ("Assunto")
    objGrid.colColuna.Add ("Anexo")
    objGrid.colColuna.Add ("Mensagem")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Selecionado.Name)
           
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        objGrid.colCampo.Add (Numero.Name)
        Parcela.left = POSICAO_FORA_TELA
        Tipo.left = POSICAO_FORA_TELA
    ElseIf giTipoTela <> TIPOTELA_EMAIL_NFE Then
        objGrid.colCampo.Add (Numero.Name)
        objGrid.colCampo.Add (Parcela.Name)
        objGrid.colCampo.Add (Tipo.Name)
    Else
        objGrid.colCampo.Add (Parcela.Name) 'Série
        objGrid.colCampo.Add (Numero.Name)
        objGrid.colCampo.Add (Tipo.Name)
    End If
    
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (Filial.Name)
    objGrid.colCampo.Add (DataVencimento.Name)
    objGrid.colCampo.Add (Valor.Name)
    
    If giTipoTela <> TIPOTELA_EMAIL_NFE Then
        If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA Then
            objGrid.colCampo.Add (Saldo.Name)
        Else
            Saldo.left = POSICAO_FORA_TELA
        End If
        
        objGrid.colCampo.Add (Atraso.Name)
    Else
        Saldo.left = POSICAO_FORA_TELA
        Atraso.left = POSICAO_FORA_TELA
    End If
    
    objGrid.colCampo.Add (Carta.Name)
    objGrid.colCampo.Add (EmailGrid.Name)
    objGrid.colCampo.Add (CCGrid.Name)
    objGrid.colCampo.Add (AssuntoGrid.Name)
    objGrid.colCampo.Add (AnexoGrid.Name)
    objGrid.colCampo.Add (MensagemGrid.Name)

    'Colunas do Grid
    iGrid_Selecionado_Col = 1
    
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        iGrid_Numero_Col = 2
        iIndice = iIndice + 2
    ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
        iGrid_Parcela_Col = 2 'Série
        iGrid_Numero_Col = 3
        iGrid_Tipo_Col = 4
    Else
        iGrid_Numero_Col = 2
        iGrid_Parcela_Col = 3
        iGrid_Tipo_Col = 4
    End If
    
    iGrid_Cliente_Col = 5 - iIndice
    iGrid_Filial_Col = 6 - iIndice
    iGrid_DataVencimento_Col = 7 - iIndice
    iGrid_Valor_Col = 8 - iIndice
    
    If giTipoTela <> TIPOTELA_EMAIL_NFE Then
        If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA Then
            iGrid_Saldo_Col = 9 - iIndice
        Else
            iIndice = iIndice + 1
        End If
        
        iGrid_Atraso_Col = 10 - iIndice
    Else
        iIndice = iIndice + 2
    End If
    iGrid_Carta_Col = 11 - iIndice
    iGrid_Email_Col = 12 - iIndice
    iGrid_CC_Col = 13 - iIndice
    iGrid_Assunto_Col = 14 - iIndice
    iGrid_Anexo_Col = 15 - iIndice
    iGrid_Mensagem_Col = 16 - iIndice

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 500 + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim lCartaAnt As Long
Dim iAtraso As Integer
Dim colCobranca As New Collection
Dim objCobranca As ClassCobrancaEmailPadrao
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable
                
    Select Case objControl.Name
    
        Case Carta.Name
        
            objControl.Enabled = True
            
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
            
                lCartaAnt = LCodigo_Extrai(Carta.Text)
                
                Carta.Clear
                
                iAtraso = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Atraso_Col))
                
                'Le os modelos válidos para o atraso em questão
                lErro = CF("CobrancaEmailPadrao_Le_ComAtraso", colCobranca, iAtraso)
                If lErro <> SUCESSO Then gError 187015
                    
                'Carrega a Combo com os Dados da Colecao
                For Each objCobranca In colCobranca
                    Carta.AddItem (objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao)
                    Carta.ItemData(Carta.NewIndex) = objCobranca.lCodigo
                Next
                    
                'Tento selecionar na Combo a Unidade anterior
                If Carta.ListCount <> 0 Then
                    For iIndice = 0 To Carta.ListCount - 1
                        If Carta.ItemData(iIndice) = lCartaAnt Then
                            Carta.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
                
            End If
            
        Case Selecionado.Name
            objControl.Enabled = True
        
        Case Else
            objControl.Enabled = False
    
    End Select
                
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 187015

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 187016)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Cliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Filial_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Filial
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Selecionado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Selecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Selecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Numero_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Numero_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Parcela_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Parcela_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Parcela_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Parcela
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Tipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Tipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Saldo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Saldo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Saldo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Saldo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Saldo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Atraso_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Atraso_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Atraso_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Atraso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Atraso
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Carta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Carta_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Carta_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Carta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Carta
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub EmailGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmailGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub EmailGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub EmailGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = EmailGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CCGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CCGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub CCGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub CCGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CCGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AssuntoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub AssuntoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub AssuntoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub AssuntoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AssuntoGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AnexoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub AnexoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub AnexoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub AnexoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AnexoGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MensagemGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MensagemGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub MensagemGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub MensagemGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = MensagemGrid
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function ValidaSelecao() As Long

Dim lErro As Long

On Error GoTo Erro_ValidaSelecao

    If giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
        'If Len(Trim(AtrasoDe.Text)) = 0 Then gError 187017
        'If Len(Trim(AtrasoAte.Text)) = 0 Then gError 187018
        
        'If StrParaInt(AtrasoDe.Text) > StrParaInt(AtrasoAte.Text) Then gError 187019
        If StrParaInt(AtrasoDe.Text) > StrParaInt(AtrasoAte.Text) And StrParaInt(AtrasoDe.Text) <> 0 And StrParaInt(AtrasoAte.Text) <> 0 Then gError 187019
    
        If Len(Trim(SaldoDe.Text)) <> 0 And Len(Trim(SaldoAte.Text)) <> 0 Then
            If StrParaDbl(SaldoDe.Text) > StrParaDbl(SaldoAte.Text) Then gError 187022
        End If
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        If StrParaInt(AtrasoDe.Text) > StrParaInt(AtrasoAte.Text) And StrParaInt(AtrasoDe.Text) <> 0 And StrParaInt(AtrasoAte.Text) <> 0 Then gError 187019
    ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Or giTipoTela = TIPOTELA_EMAIL_NFE Then
    
        If StrParaDate(DataBaixaDe.Text) <> DATA_NULA And StrParaDate(DataBaixaAte.Text) <> DATA_NULA Then
            If StrParaDate(DataBaixaDe.Text) > StrParaDate(DataBaixaAte.Text) Then gError 189428
        End If

    End If
   
    If Len(Trim(ClienteInicial.Text)) <> 0 And Len(Trim(ClienteFinal.Text)) <> 0 Then
        If LCodigo_Extrai(ClienteInicial.Text) > LCodigo_Extrai(ClienteFinal.Text) Then gError 187020
    End If
    
    If Len(Trim(TituloInic.Text)) <> 0 And Len(Trim(TituloFim.Text)) <> 0 Then
        If StrParaLong(TituloInic.Text) > StrParaLong(TituloFim.Text) Then gError 187021
    End If
    
    ValidaSelecao = SUCESSO
    
    Exit Function
    
Erro_ValidaSelecao:

    ValidaSelecao = gErr

    Select Case gErr
    
        Case 187017, 187018
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXAATRASO_NAO_PREENCHIDO", gErr)
                
        Case 187019
            Call Rotina_Erro(vbOKOnly, "ERRO_ATRASODE_MAIOR_ATRASOATE", gErr)
        
        Case 187020
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
        
        Case 187021
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", gErr)
        
        Case 187022
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            
        Case 189428
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187023)
    
    End Select

    Exit Function

End Function

Function Limpa_Tela_Email() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Email
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridItens)
        
    iAlterado = 0
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    Modelo.ListIndex = -1
        
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_SELECAO
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    Call TabStrip1_Click
    
    Call Ordenacao_Limpa(objGridItens)
    
    EmailValido.Checked = vbUnchecked
    
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
           
        '#####################################
        'Inserido por Wagner
        CategoriaClienteTodas.Value = vbChecked
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
        CategoriaClienteDe.ListIndex = -1
        CategoriaClienteAte.ListIndex = -1
        '#####################################
        
    End If
    
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        comboCobrador.ListIndex = -1
        comboCobrador.Text = ""
        ComboCobradorBordero.ListIndex = -1
        ComboCobradorBordero.Text = ""
    End If

    Limpa_Tela_Email = SUCESSO

    Exit Function

Erro_Limpa_Tela_Email:

    Limpa_Tela_Email = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187024)

    End Select

    Exit Function

End Function

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 187025

    Call Limpa_Tela_Email
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
       
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 187025

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187026)

    End Select

    Exit Sub

End Sub

Function Move_TabSelecao_Memoria(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long, iTipo As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    If giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        objCobrancaEmailSel.dSaldoAte = StrParaDbl(SaldoAte.Text)
        objCobrancaEmailSel.dSaldoDe = StrParaDbl(SaldoDe.Text)
        objCobrancaEmailSel.iAtrasoAte = StrParaInt(AtrasoAte.Text)
        objCobrancaEmailSel.iAtrasoDe = StrParaInt(AtrasoDe.Text)
        If giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
        
            objCobrancaEmailSel.iAtrasoAte = -StrParaInt(AtrasoDe.Text)
            objCobrancaEmailSel.iAtrasoDe = -StrParaInt(AtrasoAte.Text)
        
            If objCobrancaEmailSel.iAtrasoDe = 0 Then objCobrancaEmailSel.iAtrasoDe = -9999
        
'            If objCobrancaEmailSel.iAtrasoAte <> 0 Then objCobrancaEmailSel.iAtrasoAte = -objCobrancaEmailSel.iAtrasoAte
'            If objCobrancaEmailSel.iAtrasoDe <> 0 Then
'                objCobrancaEmailSel.iAtrasoDe = -objCobrancaEmailSel.iAtrasoDe
'            Else
'                If Len(Trim(AtrasoDe.Text)) = 0 Then
'                    objCobrancaEmailSel.iAtrasoDe = -9999
'                End If
'            End If
        Else
            If objCobrancaEmailSel.iAtrasoDe = 0 Then objCobrancaEmailSel.iAtrasoDe = 1
            If objCobrancaEmailSel.iAtrasoAte = 0 Then objCobrancaEmailSel.iAtrasoAte = 9999
        End If
        objCobrancaEmailSel.sCobrador = comboCobrador.Text
        objCobrancaEmailSel.iCobrador = Codigo_Extrai(ComboCobradorBordero.Text)
        objCobrancaEmailSel.lNumBordero = StrParaLong(BorderoCobrDe.Text)
    ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        objCobrancaEmailSel.dtDataBaixaAte = StrParaDate(DataBaixaAte.Text)
        objCobrancaEmailSel.dtDataBaixaDe = StrParaDate(DataBaixaDe.Text)
    ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        objCobrancaEmailSel.dtDataEmissaoAte = StrParaDate(DataBaixaAte.Text)
        objCobrancaEmailSel.dtDataEmissaoDe = StrParaDate(DataBaixaDe.Text)
        objCobrancaEmailSel.dSaldoAte = StrParaDbl(SaldoAte.Text)
        objCobrancaEmailSel.dSaldoDe = StrParaDbl(SaldoDe.Text)
    ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
        objCobrancaEmailSel.dtDataEmissaoAte = StrParaDate(DataBaixaAte.Text)
        objCobrancaEmailSel.dtDataEmissaoDe = StrParaDate(DataBaixaDe.Text)
    End If
    objCobrancaEmailSel.lClienteAte = LCodigo_Extrai(ClienteFinal.Text)
    objCobrancaEmailSel.lClienteDe = LCodigo_Extrai(ClienteInicial.Text)
    objCobrancaEmailSel.lTituloAte = StrParaLong(TituloFim.Text)
    objCobrancaEmailSel.lTituloDe = StrParaLong(TituloInic.Text)
    
    If IgnoraJaEnviados.Value = vbChecked Then
        objCobrancaEmailSel.iIgnoraJaEnviados = MARCADO
    Else
        objCobrancaEmailSel.iIgnoraJaEnviados = DESMARCADO
    End If
    
    If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
        objCobrancaEmailSel.sTipo = SCodigo_Extrai(TipoDocSeleciona.Text)
    End If
    
    objCobrancaEmailSel.lModeloForcado = LCodigo_Extrai(Modelo.Text)
    
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then

        '###########################################
        'Inserido por Wagner
        'Se a opção para todos os tipos estiver selecionada
        If CategoriaClienteTodas.Value = vbChecked Then
            objCobrancaEmailSel.sCategoria = ""
            objCobrancaEmailSel.sCategoriaDe = ""
            objCobrancaEmailSel.sCategoriaAte = ""
        Else
            If CategoriaCliente.Text = "" Then gError 198622
            objCobrancaEmailSel.sCategoria = CategoriaCliente.Text
            objCobrancaEmailSel.sCategoriaDe = CategoriaClienteDe.Text
            objCobrancaEmailSel.sCategoriaAte = CategoriaClienteAte.Text
        End If
        
        If objCobrancaEmailSel.sCategoriaDe > objCobrancaEmailSel.sCategoriaAte Then gError 198623
        '###########################################

    End If
    
    If EmailValido.Value = vbChecked Then
        objCobrancaEmailSel.iSoComEmailValido = MARCADO
    Else
        objCobrancaEmailSel.iSoComEmailValido = DESMARCADO
    End If
    
    Select Case giTipoTela
    
        Case TIPOTELA_EMAIL_AGRADECIMENTO
            iTipo = TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO
            
        Case TIPOTELA_EMAIL_COBRANCA_FATURA
            iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA
            
        Case TIPOTELA_EMAIL_NFE
            iTipo = TIPO_COBRANCAEMAILPADRAO_NFE
            
        Case TIPOTELA_EMAIL_AVISO_PAGTO_CP
            iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP
    
        Case TIPOTELA_EMAIL_COBRANCA
            iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA
    
        Case TIPOTELA_EMAIL_AVISO_COBRANCA
            iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO
       
    End Select
    
    objCobrancaEmailSel.iTipoTela = iTipo
    
    lErro = CF2(Me, "CE_Move_TabSelecao_Memoria", objCobrancaEmailSel)
    If lErro <> SUCESSO Then gError 196970
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr
    
        Case 196970
        
        '###################################################
        'Inserido por Wagner
        Case 198622
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
            
        Case 198623
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ITEM_INICIAL_MAIOR", gErr)
        '###################################################

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187027)

    End Select

    Exit Function

End Function

Function CE_Move_TabSelecao_Memoria(ByVal objCT As Object, ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
    CE_Move_TabSelecao_Memoria = SUCESSO
End Function

Public Sub BotaoMarcarTodos_Click()
'Marca todos os itens do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = MARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridItens)

End Sub

Public Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os itens do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = DESMARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridItens)

End Sub

Private Function Preenche_GridItens(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
'Preenche o Grid Pedido com os dados de colItens

Dim lErro As Long

On Error GoTo Erro_Preenche_GridItens

    Call Grid_Limpa(objGridItens)
    
    Call Ordenacao_Limpa(objGridItens)
    
    Set gcolClientes = New Collection
    Set gcolFiliaisCli = New Collection
    Set gcolParcelas = New Collection
    Set gcolTitulos = New Collection
    Set gcolModeloLinha = New Collection
    Set gcolEndereco = New Collection
    
    'Aumenta o número de linhas do grid se necessário
    If objCobrancaEmailSel.colParcelas.Count >= objGridItens.objGrid.Rows Then
        Call Refaz_Grid(objGridItens, objCobrancaEmailSel.colParcelas.Count)
    End If

    'Le os modelos válidos para o atraso em questão
    If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
        lErro = Preenche_GridItens_CP1(objCobrancaEmailSel)
    ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        lErro = Preenche_GridItens_CP2(objCobrancaEmailSel)
    ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
        lErro = Preenche_GridItens_FAT(objCobrancaEmailSel)
    Else
        lErro = Preenche_GridItens_CR(objCobrancaEmailSel)
    End If
    If lErro <> SUCESSO Then gError 196972
    
    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(2).Visible = True
    'Torna Frame atual visivel
    Frame1(1).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = 2

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr
    
        Case 196972

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196973)

    End Select

    Exit Function

End Function

Private Function Preenche_GridItens_CR(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
'Preenche o Grid Pedido com os dados de colItens

Dim lErro As Long
Dim iLinha As Integer
Dim objParcelaRec As ClassParcelaReceber
Dim objTituloRec As ClassTituloReceber
Dim objCobranca As ClassCobrancaEmailPadrao
Dim colCliente As New Collection
Dim bAchou As Boolean
Dim iAtraso As Integer
Dim lNumIntTituloAnt As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim colFiliais As New Collection
Dim colCobranca As Collection
Dim objEndereco As ClassEndereco
Dim colEndereco As New Collection
Dim iPosicao As Integer
Dim objTela As Object, sEmail As String

On Error GoTo Erro_Preenche_GridItens_CR

    Set objTela = Me

    iLinha = 0
    
    'Percorre todas as OP da Coleção
    For Each objParcelaRec In objCobrancaEmailSel.colParcelas

        If lNumIntTituloAnt <> objParcelaRec.lNumIntTitulo Then

            Set objTituloRec = New ClassTituloReceber
            
            objTituloRec.lNumIntDoc = objParcelaRec.lNumIntTitulo
            
            lErro = CF("TituloReceber_Le", objTituloRec)
            If lErro <> SUCESSO And lErro <> 26061 Then gError 187028
            
            If lErro <> SUCESSO Then
            
                lErro = CF("TituloReceberBaixado_Le", objTituloRec)
                If lErro <> SUCESSO And lErro <> 56570 Then gError 187029
                
            End If
            
            lNumIntTituloAnt = objParcelaRec.lNumIntTitulo
            
        End If
        
        bAchou = False
        For Each objCliente In colCliente
            If objCliente.lCodigo = objTituloRec.lCliente Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objCliente = New ClassCliente
            
            objCliente.lCodigo = objTituloRec.lCliente
        
            'le o cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 187030
                    
            colCliente.Add objCliente
                    
        End If
        
        bAchou = False
        iPosicao = 0
        For Each objFilialCliente In colFiliais
            iPosicao = iPosicao + 1
            If objTituloRec.lCliente = objFilialCliente.lCodCliente And objFilialCliente.iCodFilial = objTituloRec.iFilial Then
                bAchou = True
                Set objEndereco = colEndereco.Item(iPosicao)
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objFilialCliente = New ClassFilialCliente
            Set objEndereco = New ClassEndereco
            
            objFilialCliente.lCodCliente = objTituloRec.lCliente
            objFilialCliente.iCodFilial = objTituloRec.iFilial
    
            'le o nome reduzido da filial  cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 187031
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 187032

            lErro = CF("Endereco_Trata_Customizacao", objEndereco)
            If lErro <> SUCESSO Then gError 196232

            colEndereco.Add objEndereco
            colFiliais.Add objFilialCliente
            
        End If
        
        If objCobrancaEmailSel.iSoComEmailValido = DESMARCADO Or (InStr(1, objEndereco.sEmail, "@") <> 0 And InStr(1, objEndereco.sEmail, ".") <> 0) Then
    
            Set colCobranca = New Collection
    
            iLinha = iLinha + 1
            
            iAtraso = DateDiff("d", objParcelaRec.dtDataVencimentoReal, gdtDataHoje)
            
            'Le os modelos válidos para o atraso em questão
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
                lErro = CF("CobrancaEmailPadrao_Le_ComAtraso", colCobranca, iAtraso)
            ElseIf giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Then
                lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO)
            ElseIf giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
                lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_AVISO)
            End If
            If lErro <> SUCESSO Then gError 187033
            
            'Força o modelo de email que estará como padrão
            If objCobrancaEmailSel.lModeloForcado <> 0 Then
            
                bAchou = False
                For Each objCobranca In colCobranca
                    If objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                
                    Set objCobranca = New ClassCobrancaEmailPadrao
                
                    objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado
                    
                    lErro = CF("CobrancaEmailPadrao_Le", objCobranca)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192213
                    
                    colCobranca.Add objCobranca
                    
                End If
            
            End If
            
            'Se não tiver algum modelo
            If colCobranca.Count = 0 Then
                Set objCobranca = New ClassCobrancaEmailPadrao
            Else
                If objCobrancaEmailSel.lModeloForcado = 0 Then
                   'Se tiver escolhe o primeiro e calcula os valores das regras
                    Set objCobranca = colCobranca.Item(1)
                End If
            End If
    
            If objCobrancaEmailSel.iFaturamento = DESMARCADO Then
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(DESMARCADO)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objTituloRec.lNumTitulo)
            GridItens.TextMatrix(iLinha, iGrid_Parcela_Col) = CStr(objParcelaRec.iNumParcela)
            GridItens.TextMatrix(iLinha, iGrid_Tipo_Col) = objTituloRec.sSiglaDocumento
            GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
            GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
            GridItens.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objParcelaRec.dtDataVencimentoReal, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objParcelaRec.dValor, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objParcelaRec.dSaldo, "STANDARD")
            
            If giTipoTela <> TIPOTELA_EMAIL_AGRADECIMENTO Then
                If iAtraso < 0 Then iAtraso = -iAtraso
                GridItens.TextMatrix(iLinha, iGrid_Atraso_Col) = CStr(iAtraso)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Atraso_Col) = Format(objParcelaRec.dtDataUltimaBaixa, "dd/mm/yyyy")
            End If
            
            If objCobranca.lCodigo <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Carta_Col) = objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao
            End If
            
            sEmail = objEndereco.sEmail
            If Len(Trim(objEndereco.sEmail2)) > 0 Then
                If sEmail <> "" Then sEmail = sEmail & ";"
                sEmail = sEmail & objEndereco.sEmail2
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Email_Col) = sEmail
            
            gcolClientes.Add objCliente
            gcolFiliaisCli.Add objFilialCliente
            gcolParcelas.Add objParcelaRec
            gcolTitulos.Add objTituloRec
            gcolModeloLinha.Add objCobranca
            gcolEndereco.Add objEndereco
        
        End If

    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridItens.iLinhasExistentes = iLinha
        
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        giLinhaAtual = iLinha
    
        Set objCobranca = gcolModeloLinha.Item(iLinha)
    
        lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objCobranca)
        If lErro <> SUCESSO Then gError 187034
    
        GridItens.TextMatrix(iLinha, iGrid_CC_Col) = objCobranca.sCCValor
        GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = objCobranca.sAssuntoValor
        GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = objCobranca.sMensagemValor
        GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
      
        If Len(Trim(objCobranca.sMensagemValor)) = 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem automática confome modelo '" & objCobranca.sModelo & "'>"
        End If

    Next

    Preenche_GridItens_CR = SUCESSO

    Exit Function

Erro_Preenche_GridItens_CR:

    Preenche_GridItens_CR = gErr

    Select Case gErr
    
        Case 187028 To 187034, 192213, 196232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187035)

    End Select

    Exit Function

End Function

Private Function Preenche_GridItens_FAT(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
'Preenche o Grid Pedido com os dados de colItens

Dim lErro As Long
Dim iLinha As Integer
Dim objNF As ClassNFiscal
Dim objCobranca As ClassCobrancaEmailPadrao
Dim colCliente As New Collection
Dim bAchou As Boolean
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim colFornecedor As New Collection
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colFiliaisForn As New Collection
Dim colFiliaisCli As New Collection
Dim colCobranca As Collection
Dim objEndereco As ClassEndereco
Dim colEnderecoCli As New Collection
Dim colEnderecoForn As New Collection
Dim iPosicao As Integer
Dim objTela As Object
Dim objTipoDocInfo As ClassTipoDocInfo
Dim sEmail As String
Dim objTransp As ClassTransportadora
Dim objEndTransp As ClassEndereco
Dim sDanfe As String

On Error GoTo Erro_Preenche_GridItens_FAT

    Set objTela = Me

    iLinha = 0
    
    'Percorre todas as OP da Coleção
    For Each objNF In objCobrancaEmailSel.colParcelas
    
        If objNF.lCliente <> 0 Then
        
            bAchou = False
            For Each objCliente In colCliente
                If objCliente.lCodigo = objNF.lCliente Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objCliente = New ClassCliente
                
                objCliente.lCodigo = objNF.lCliente
            
                'le o cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 187030
                        
                colCliente.Add objCliente
                        
            End If
            
            bAchou = False
            iPosicao = 0
            For Each objFilialCliente In colFiliaisCli
                iPosicao = iPosicao + 1
                If objNF.lCliente = objFilialCliente.lCodCliente And objFilialCliente.iCodFilial = objNF.iFilialCli Then
                    bAchou = True
                    Set objEndereco = colEnderecoCli.Item(iPosicao)
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objFilialCliente = New ClassFilialCliente
                Set objEndereco = New ClassEndereco
                
                objFilialCliente.lCodCliente = objNF.lCliente
                objFilialCliente.iCodFilial = objNF.iFilialCli
        
                'le o nome reduzido da filial  cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 187031
                
                objEndereco.lCodigo = objFilialCliente.lEndereco
            
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 187032
    
                lErro = CF("Endereco_Trata_Customizacao", objEndereco)
                If lErro <> SUCESSO Then gError 196232
    
                colEnderecoCli.Add objEndereco
                colFiliaisCli.Add objFilialCliente
                
            End If
        
        Else 'fornecedores
        
            bAchou = False
            For Each objFornecedor In colFornecedor
                If objFornecedor.lCodigo = objNF.lFornecedor Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objFornecedor = New ClassFornecedor
                
                objFornecedor.lCodigo = objNF.lFornecedor
            
                'le o Fornecedor
                lErro = CF("Fornecedor_Le", objFornecedor)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 187030
                        
                colFornecedor.Add objFornecedor
                        
            End If
            
            bAchou = False
            iPosicao = 0
            For Each objFilialFornecedor In colFiliaisForn
                iPosicao = iPosicao + 1
                If objNF.lFornecedor = objFilialFornecedor.lCodFornecedor And objFilialFornecedor.iCodFilial = objNF.iFilialForn Then
                    bAchou = True
                    Set objEndereco = colEnderecoForn.Item(iPosicao)
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objFilialFornecedor = New ClassFilialFornecedor
                Set objEndereco = New ClassEndereco
                
                objFilialFornecedor.lCodFornecedor = objNF.lFornecedor
                objFilialFornecedor.iCodFilial = objNF.iFilialForn
        
                'le o nome reduzido da filial  Fornecedor
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 187031
                
                objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 187032
    
                lErro = CF("Endereco_Trata_Customizacao", objEndereco)
                If lErro <> SUCESSO Then gError 196232
    
                colEnderecoForn.Add objEndereco
                colFiliaisForn.Add objFilialFornecedor
                
            End If
        
        End If

        If objCobrancaEmailSel.iSoComEmailValido = DESMARCADO Or (InStr(1, objEndereco.sEmail, "@") <> 0 And InStr(1, objEndereco.sEmail, ".") <> 0) Then
    
            Set colCobranca = New Collection
    
            iLinha = iLinha + 1
            
            'Le os modelos válidos
            lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_NFE)
            If lErro <> SUCESSO Then gError 187033
            
            'Força o modelo de email que estará como padrão
            If objCobrancaEmailSel.lModeloForcado <> 0 Then
            
                bAchou = False
                For Each objCobranca In colCobranca
                    If objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado Then
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If Not bAchou Then
                
                    Set objCobranca = New ClassCobrancaEmailPadrao
                
                    objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado
                    
                    lErro = CF("CobrancaEmailPadrao_Le", objCobranca)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192213
                    
                    colCobranca.Add objCobranca
                    
                End If
            
            End If
            
            'Se não tiver algum modelo
            If colCobranca.Count = 0 Then
                Set objCobranca = New ClassCobrancaEmailPadrao
            Else
                If objCobrancaEmailSel.lModeloForcado = 0 Then
                   'Se tiver escolhe o primeiro e calcula os valores das regras
                    Set objCobranca = colCobranca.Item(1)
                End If
            End If
    
            If objCobrancaEmailSel.iFaturamento = DESMARCADO Then
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(DESMARCADO)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
            End If
            
            Set objTipoDocInfo = New ClassTipoDocInfo
            objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
            
            'Lê o tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 192213
    
            GridItens.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objNF.lNumNotaFiscal)
            GridItens.TextMatrix(iLinha, iGrid_Parcela_Col) = objNF.sSerie
            GridItens.TextMatrix(iLinha, iGrid_Tipo_Col) = CStr(objNF.iTipoNFiscal) & SEPARADOR & objTipoDocInfo.sDescricao
            
            If objNF.lCliente <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
                GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
            Else
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objFornecedor.lCodigo) & SEPARADOR & objFornecedor.sNomeReduzido
                GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objNF.dtDataEmissao, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objNF.dValorTotal, "STANDARD")
            
            If objCobranca.lCodigo <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Carta_Col) = objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao
            End If
            
            sEmail = ""
            If cCliente.Value = vbChecked Then
                sEmail = objEndereco.sEmail
                If Len(Trim(objEndereco.sEmail2)) > 0 Then
                    If sEmail <> "" Then sEmail = sEmail & ";"
                    sEmail = sEmail & objEndereco.sEmail2
                End If
            End If
            If cTransportadora.Value = vbChecked Then
            
                If objNF.iCodTransportadora <> 0 Then
                
                    Set objTransp = New ClassTransportadora
                    Set objEndTransp = New ClassEndereco
                    
                    objTransp.iCodigo = objNF.iCodTransportadora
                    
                    lErro = CF("Transportadora_Le", objTransp)
                    If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
            
                    'Carrega lEndereco em objTransportadora
                    objEndTransp.lCodigo = objTransp.lEndereco
                
                    'Lê o endereço a partir do Código
                    lErro = CF("Endereco_Le", objEndTransp)
                    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                
                    If Len(Trim(objEndTransp.sEmail)) > 0 Then
                        If sEmail <> "" Then sEmail = sEmail & ";"
                        sEmail = sEmail & objEndTransp.sEmail
                    End If
                    
                    If Len(Trim(objEndTransp.sEmail2)) > 0 Then
                        If sEmail <> "" Then sEmail = sEmail & ";"
                        sEmail = sEmail & objEndTransp.sEmail2
                    End If
                    
                End If
                
                If objNF.iCodTranspRedesp <> 0 And objNF.iCodTransportadora <> objNF.iCodTranspRedesp Then
                
                    Set objTransp = New ClassTransportadora
                    Set objEndTransp = New ClassEndereco
                    
                    objTransp.iCodigo = objNF.iCodTranspRedesp
                    
                    lErro = CF("Transportadora_Le", objTransp)
                    If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
            
                    'Carrega lEndereco em objTransportadora
                    objEndTransp.lCodigo = objTransp.lEndereco
                
                    'Lê o endereço a partir do Código
                    lErro = CF("Endereco_Le", objEndTransp)
                    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                
                    If Len(Trim(objEndTransp.sEmail)) > 0 Then
                        If sEmail <> "" Then sEmail = sEmail & ";"
                        sEmail = sEmail & objEndTransp.sEmail
                    End If
                    
                    If Len(Trim(objEndTransp.sEmail2)) > 0 Then
                        If sEmail <> "" Then sEmail = sEmail & ";"
                        sEmail = sEmail & objEndTransp.sEmail2
                    End If
                    
                End If
            
            End If
            If cFixo.Value = vbChecked Then
                If Len(Trim(emailFixo.Text)) > 0 Then
                    If sEmail <> "" Then sEmail = sEmail & ";" & emailFixo.Text
                End If
            End If
            GridItens.TextMatrix(iLinha, iGrid_Email_Col) = sEmail 'objEndereco.sEmail
            
            If objNF.lCliente <> 0 Then
                gcolClientes.Add objCliente
                gcolFiliaisCli.Add objFilialCliente
            Else
                gcolClientes.Add objFornecedor
                gcolFiliaisCli.Add objFilialFornecedor
            End If
            gcolParcelas.Add objNF
            gcolTitulos.Add objNF
            gcolModeloLinha.Add objCobranca
            gcolEndereco.Add objEndereco
        
        End If
    
    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridItens.iLinhasExistentes = iLinha
        
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        giLinhaAtual = iLinha
    
        Set objCobranca = gcolModeloLinha.Item(iLinha)
    
        lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objCobranca)
        If lErro <> SUCESSO Then gError 187034
    
        GridItens.TextMatrix(iLinha, iGrid_CC_Col) = objCobranca.sCCValor
        GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = objCobranca.sAssuntoValor
        GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = objCobranca.sMensagemValor
        
        If optIncluiDanfe.Value = vbChecked Then
        
            Set objNF = gcolTitulos.Item(iLinha)
        
            lErro = CF("NFe_Obtem_Nome_Danfe", objNF, sDanfe)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Verifica_Existencia_Arquivo2(sDanfe, True)
            If lErro = SUCESSO Then
                GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor & ";" & sDanfe
            Else
                GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
            End If
        
        Else
            GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
        End If
        
        If Len(Trim(objCobranca.sMensagemValor)) = 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem automática confome modelo '" & objCobranca.sModelo & "'>"
        End If

    Next

    Preenche_GridItens_FAT = SUCESSO

    Exit Function

Erro_Preenche_GridItens_FAT:

    Preenche_GridItens_FAT = gErr

    Select Case gErr
    
        Case 187028 To 187034, 192213, 196232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187035)

    End Select

    Exit Function

End Function

Private Function Preenche_GridItens_CP1(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
'Preenche o Grid Pedido com os dados de colItens

Dim lErro As Long
Dim iLinha As Integer
Dim objNFsPag As ClassNFsPag
Dim objCobranca As ClassCobrancaEmailPadrao
Dim colForn As New Collection
Dim bAchou As Boolean
Dim iAtraso As Integer
Dim lNumIntTituloAnt As Long
Dim objForn As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor
Dim colFiliais As New Collection
Dim colCobranca As Collection
Dim objEndereco As ClassEndereco
Dim colEndereco As New Collection
Dim iPosicao As Integer
Dim objTela As Object, sEmail As String

On Error GoTo Erro_Preenche_GridItens_CP1

    Set objTela = Me

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objNFsPag In objCobrancaEmailSel.colParcelas
           
        bAchou = False
        For Each objForn In colForn
            If objForn.lCodigo = objNFsPag.lFornecedor Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objForn = New ClassFornecedor
            
            objForn.lCodigo = objNFsPag.lFornecedor
        
            'le o cliente
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 And lErro <> 12732 Then gError 187030
                    
            colForn.Add objForn
                    
        End If
        
        bAchou = False
        iPosicao = 0
        For Each objFilialForn In colFiliais
            iPosicao = iPosicao + 1
            If objNFsPag.lFornecedor = objFilialForn.lCodFornecedor And objFilialForn.iCodFilial = objNFsPag.iFilial Then
                bAchou = True
                Set objEndereco = colEndereco.Item(iPosicao)
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objFilialForn = New ClassFilialFornecedor
            Set objEndereco = New ClassEndereco
            
            objFilialForn.lCodFornecedor = objNFsPag.lFornecedor
            objFilialForn.iCodFilial = objNFsPag.iFilial
    
            'le o nome reduzido da filial  cliente
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 187031
            
            objEndereco.lCodigo = objFilialForn.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 187032

            colEndereco.Add objEndereco
            colFiliais.Add objFilialForn
            
        End If
        
        If objCobrancaEmailSel.iSoComEmailValido = DESMARCADO Or (InStr(1, objEndereco.sEmail, "@") <> 0 And InStr(1, objEndereco.sEmail, ".") <> 0) Then
        
            Set colCobranca = New Collection
    
            iLinha = iLinha + 1
        
            'Le os modelos válidos para o atraso em questão
            lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA)
            If lErro <> SUCESSO Then gError 187033
            
            'Força o modelo de email que estará como padrão
            If objCobrancaEmailSel.lModeloForcado <> 0 Then
            
                bAchou = False
                For Each objCobranca In colCobranca
                    If objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                
                    Set objCobranca = New ClassCobrancaEmailPadrao
                
                    objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado
                    
                    lErro = CF("CobrancaEmailPadrao_Le", objCobranca)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192213
                    
                    colCobranca.Add objCobranca
                    
                End If
            
            End If
            
            'Se não tiver algum modelo
            If colCobranca.Count = 0 Then
                Set objCobranca = New ClassCobrancaEmailPadrao
            Else
                If objCobrancaEmailSel.lModeloForcado = 0 Then
                   'Se tiver escolhe o primeiro e calcula os valores das regras
                    Set objCobranca = colCobranca.Item(1)
                End If
            End If
    
            If objCobrancaEmailSel.iFaturamento = DESMARCADO Then
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(DESMARCADO)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objNFsPag.lNumNotaFiscal)
            GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
            GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialForn.iCodFilial) & SEPARADOR & objFilialForn.sNome
            GridItens.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objNFsPag.dtDataVencimento, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objNFsPag.dValorTotal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_Atraso_Col) = Format(objNFsPag.dtDataEmissao, "dd/mm/yyyy")
            
            If objCobranca.lCodigo <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Carta_Col) = objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao
            End If
            
            sEmail = objEndereco.sEmail
            If Len(Trim(objEndereco.sEmail2)) > 0 Then
                If sEmail <> "" Then sEmail = sEmail & ";"
                sEmail = sEmail & objEndereco.sEmail2
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Email_Col) = sEmail
            
            gcolClientes.Add objForn
            gcolFiliaisCli.Add objFilialForn
            gcolParcelas.Add objNFsPag
            gcolTitulos.Add objNFsPag
            gcolModeloLinha.Add objCobranca
            gcolEndereco.Add objEndereco
            
        End If

    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridItens.iLinhasExistentes = iLinha
        
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        giLinhaAtual = iLinha
    
        Set objCobranca = gcolModeloLinha.Item(iLinha)
    
        lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objCobranca)
        If lErro <> SUCESSO Then gError 187034
    
        GridItens.TextMatrix(iLinha, iGrid_CC_Col) = objCobranca.sCCValor
        GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = objCobranca.sAssuntoValor
        GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = objCobranca.sMensagemValor
        GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
        
        If Len(Trim(objCobranca.sMensagemValor)) = 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem automática confome modelo '" & objCobranca.sModelo & "'>"
        End If

    Next

    Preenche_GridItens_CP1 = SUCESSO

    Exit Function

Erro_Preenche_GridItens_CP1:

    Preenche_GridItens_CP1 = gErr

    Select Case gErr
    
        Case 187028 To 187034, 192213, 196232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187035)

    End Select

    Exit Function

End Function

Private Function Preenche_GridItens_CP2(ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long
'Preenche o Grid Pedido com os dados de colItens

Dim lErro As Long
Dim iLinha As Integer
Dim objTitPag As ClassTituloPagar
Dim objParcelaPag As ClassParcelaPagar
Dim objCobranca As ClassCobrancaEmailPadrao
Dim colForn As New Collection
Dim bAchou As Boolean
Dim iAtraso As Integer
Dim lNumIntTituloAnt As Long
Dim objForn As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor
Dim colFiliais As New Collection
Dim colCobranca As Collection
Dim objEndereco As ClassEndereco
Dim colEndereco As New Collection
Dim iPosicao As Integer
Dim objTela As Object, sEmail As String

On Error GoTo Erro_Preenche_GridItens_CP2

    Set objTela = Me

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objParcelaPag In objCobrancaEmailSel.colParcelas
    
        Set colCobranca = New Collection

        If lNumIntTituloAnt <> objParcelaPag.lNumIntTitulo Then

            Set objTitPag = New ClassTituloPagar
            
            objTitPag.lNumIntDoc = objParcelaPag.lNumIntTitulo
            
            lErro = CF("TituloPagar_Le_Todos", objTitPag)
            If lErro <> SUCESSO Then gError 187028
            
            lNumIntTituloAnt = objParcelaPag.lNumIntTitulo
            
        End If
        
        bAchou = False
        For Each objForn In colForn
            If objForn.lCodigo = objTitPag.lFornecedor Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objForn = New ClassFornecedor
            
            objForn.lCodigo = objTitPag.lFornecedor
        
            'le o cliente
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 And lErro <> 12732 Then gError 187030
                    
            colForn.Add objForn
                    
        End If
        
        bAchou = False
        iPosicao = 0
        For Each objFilialForn In colFiliais
            iPosicao = iPosicao + 1
            If objTitPag.lFornecedor = objFilialForn.lCodFornecedor And objFilialForn.iCodFilial = objTitPag.iFilial Then
                bAchou = True
                Set objEndereco = colEndereco.Item(iPosicao)
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objFilialForn = New ClassFilialFornecedor
            Set objEndereco = New ClassEndereco
            
            objFilialForn.lCodFornecedor = objTitPag.lFornecedor
            objFilialForn.iCodFilial = objTitPag.iFilial
    
            'le o nome reduzido da filial  cliente
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 187031
            
            objEndereco.lCodigo = objFilialForn.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 187032

            colEndereco.Add objEndereco
            colFiliais.Add objFilialForn
            
        End If
        
        If objCobrancaEmailSel.iSoComEmailValido = DESMARCADO Or (InStr(1, objEndereco.sEmail, "@") <> 0 And InStr(1, objEndereco.sEmail, ".") <> 0) Then
                    
            iLinha = iLinha + 1
                    
            'Le os modelos válidos para o atraso em questão
            lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP)
            If lErro <> SUCESSO Then gError 187033
            
            'Força o modelo de email que estará como padrão
            If objCobrancaEmailSel.lModeloForcado <> 0 Then
            
                bAchou = False
                For Each objCobranca In colCobranca
                    If objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                
                    Set objCobranca = New ClassCobrancaEmailPadrao
                
                    objCobranca.lCodigo = objCobrancaEmailSel.lModeloForcado
                    
                    lErro = CF("CobrancaEmailPadrao_Le", objCobranca)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192213
                    
                    colCobranca.Add objCobranca
                    
                End If
            
            End If
            
            'Se não tiver algum modelo
            If colCobranca.Count = 0 Then
                Set objCobranca = New ClassCobrancaEmailPadrao
            Else
                If objCobrancaEmailSel.lModeloForcado = 0 Then
                   'Se tiver escolhe o primeiro e calcula os valores das regras
                    Set objCobranca = colCobranca.Item(1)
                End If
            End If
    
            If objCobrancaEmailSel.iFaturamento = DESMARCADO Then
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(DESMARCADO)
            Else
                GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objTitPag.lNumTitulo)
            GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
            GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = CStr(objFilialForn.iCodFilial) & SEPARADOR & objFilialForn.sNome
            GridItens.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objParcelaPag.dtDataVencimento, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objTitPag.dValorTotal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_Atraso_Col) = Format(objParcelaPag.dtDataUltimaBaixa, "dd/mm/yyyy")
    
            GridItens.TextMatrix(iLinha, iGrid_Parcela_Col) = CStr(objParcelaPag.iNumParcela)
            GridItens.TextMatrix(iLinha, iGrid_Tipo_Col) = objTitPag.sSiglaDocumento
            GridItens.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objParcelaPag.dSaldo, "STANDARD")
            
            If objCobranca.lCodigo <> 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Carta_Col) = objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao
            End If
            
            sEmail = objEndereco.sEmail
            If Len(Trim(objEndereco.sEmail2)) > 0 Then
                If sEmail <> "" Then sEmail = sEmail & ";"
                sEmail = sEmail & objEndereco.sEmail2
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Email_Col) = sEmail
            
            gcolClientes.Add objForn
            gcolFiliaisCli.Add objFilialForn
            gcolParcelas.Add objParcelaPag
            gcolTitulos.Add objTitPag
            gcolModeloLinha.Add objCobranca
            gcolEndereco.Add objEndereco
            
        End If

    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridItens.iLinhasExistentes = iLinha
        
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        giLinhaAtual = iLinha
    
        Set objCobranca = gcolModeloLinha.Item(iLinha)
    
        lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objCobranca)
        If lErro <> SUCESSO Then gError 187034
    
        GridItens.TextMatrix(iLinha, iGrid_CC_Col) = objCobranca.sCCValor
        GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = objCobranca.sAssuntoValor
        GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = objCobranca.sMensagemValor
        GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
        
        If Len(Trim(objCobranca.sMensagemValor)) = 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem automática confome modelo '" & objCobranca.sModelo & "'>"
        End If

    Next

    Preenche_GridItens_CP2 = SUCESSO

    Exit Function

Erro_Preenche_GridItens_CP2:

    Preenche_GridItens_CP2 = gErr

    Select Case gErr
    
        Case 187028 To 187034, 192213, 196232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187035)

    End Select

    Exit Function

End Function

Private Function Mostra_Dados_Email(ByVal iLinha As Integer) As Long

Dim objCobranca As ClassCobrancaEmailPadrao

    If iLinha <> 0 Then
        Cc.Text = GridItens.TextMatrix(iLinha, iGrid_CC_Col)
        Assunto.Text = GridItens.TextMatrix(iLinha, iGrid_Assunto_Col)
        Anexo.Text = GridItens.TextMatrix(iLinha, iGrid_Anexo_Col)
        Mensagem.Text = GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col)
        Email.Text = GridItens.TextMatrix(iLinha, iGrid_Email_Col)
        Mensagem.Enabled = False
    End If
    
    Mostra_Dados_Email = SUCESSO

End Function

Private Function Recolhe_Dados_Email(ByVal iLinha As Integer) As Long

    If iLinha <> 0 Then
        GridItens.TextMatrix(iLinha, iGrid_CC_Col) = Cc.Text
        GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = Assunto.Text
        GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = Mensagem.Text
        GridItens.TextMatrix(iLinha, iGrid_Email_Col) = Email.Text
        GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = Anexo.Text
    End If
    
    Recolhe_Dados_Email = SUCESSO

End Function

Public Sub ClienteFinal_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub ClienteInicial_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub SaldoAte_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub SaldoDe_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub TipoDocApenas_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    TipoDocSeleciona.Enabled = True
End Sub

Public Sub TipoDocSeleciona_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub TipoDocTodos_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    'Desabilita a combo para a seleção
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção
    TipoDocSeleciona.ListIndex = COMBO_INDICE
    
End Sub

Public Sub TituloFim_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub TituloInic_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Private Function Preenche_Dados_Carta(ByVal lCarta As Long, ByVal lCartaAnt As Long, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objCobranca As New ClassCobrancaEmailPadrao
Dim objTela As Object

On Error GoTo Erro_Preenche_Dados_Carta

    If iLinha <> 0 And lCarta <> 0 Then
    
        'Se trocou o tipo de carta
        If lCarta <> lCartaAnt Then

            objCobranca.lCodigo = lCarta
        
            lErro = CF("CobrancaEmailPadrao_Le", objCobranca)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187035
            
            gcolModeloLinha.Item(iLinha).iAtrasoAte = objCobranca.iAtrasoAte
            gcolModeloLinha.Item(iLinha).iAtrasoDe = objCobranca.iAtrasoDe
            gcolModeloLinha.Item(iLinha).lCodigo = objCobranca.lCodigo
            gcolModeloLinha.Item(iLinha).lNumIntDoc = objCobranca.lNumIntDoc
            gcolModeloLinha.Item(iLinha).sAssunto = objCobranca.sAssunto
            gcolModeloLinha.Item(iLinha).sAssuntoValor = objCobranca.sAssuntoValor
            gcolModeloLinha.Item(iLinha).sCC = objCobranca.sCC
            gcolModeloLinha.Item(iLinha).sCCValor = objCobranca.sCCValor
            gcolModeloLinha.Item(iLinha).sDescricao = objCobranca.sDescricao
            gcolModeloLinha.Item(iLinha).sMensagem = objCobranca.sMensagem
            gcolModeloLinha.Item(iLinha).sMensagemValor = objCobranca.sMensagemValor
            gcolModeloLinha.Item(iLinha).sModelo = objCobranca.sModelo
            gcolModeloLinha.Item(iLinha).sAnexo = objCobranca.sAnexo
            gcolModeloLinha.Item(iLinha).iTipo = objCobranca.iTipo
            gcolModeloLinha.Item(iLinha).sDe = objCobranca.sDe
            gcolModeloLinha.Item(iLinha).sNomeExibicao = objCobranca.sNomeExibicao
            
            Set objTela = Me
            
            giLinhaAtual = iLinha
        
            Set objCobranca = gcolModeloLinha.Item(iLinha)
        
            lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objCobranca)
            If lErro <> SUCESSO Then gError 187036
        
            GridItens.TextMatrix(iLinha, iGrid_CC_Col) = objCobranca.sCCValor
            GridItens.TextMatrix(iLinha, iGrid_Assunto_Col) = objCobranca.sAssuntoValor
            GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = objCobranca.sMensagemValor
            GridItens.TextMatrix(iLinha, iGrid_Anexo_Col) = objCobranca.sAnexoValor
            
            If Len(Trim(objCobranca.sMensagemValor)) = 0 Then
                GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem automática confome modelo '" & objCobranca.sModelo & "'>"
            End If
        
        End If
                
    End If
    
    Preenche_Dados_Carta = SUCESSO

    Exit Function

Erro_Preenche_Dados_Carta:

    Preenche_Dados_Carta = gErr

    Select Case gErr
    
        Case 187035, 187036

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187037)

    End Select

    Exit Function

End Function

Public Sub BotaoEmail_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim sMsgAux As String
Dim objCobrEmail As ClassCobrancaEmailPadrao
Dim iGeraRelac As Integer
Dim objEnvioEmail As ClassEnvioDeEmail
Dim colEnvioEmail As New Collection
Dim iNumParcelas As Integer
Dim lNumTitulo As Long
Dim iNumParcela As Integer
Dim lCliente As Long
Dim lNumIntParcela As Long
Dim sNomeArqParam As String
Dim vValor As Variant
Dim iFilialCliRelac As Integer
Dim objFilialEmpresa As AdmFiliais

On Error GoTo Erro_BotaoEmail_Click
    
    'GL_objMDIForm.MousePointer = vbHourglass
    
    Call Recolhe_Dados_Email(iLinhaAnt)
    
    iNumParcelas = 0
    For iLinha = 1 To objGridItens.iLinhasExistentes
                   
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
       
            iNumParcelas = iNumParcelas + 1
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Email_Col))) = 0 Then gError 187038
    
            If InStr(1, GridItens.TextMatrix(iLinha, iGrid_Email_Col), "@") = 0 Then gError 187038
    
            lErro = Verifica_Existencia_Arquivo(GridItens.TextMatrix(iLinha, iGrid_Anexo_Col))
            If lErro <> SUCESSO Then gError 189413
    
        End If
    
    Next
    
    If iNumParcelas = 0 Then gError 187118
    
    Set gcolParcelasEnv = New Collection
    Set gcolTitulosEnv = New Collection
    Set gcolClientesEnv = New Collection
    Set gcolFiliaisCliEnv = New Collection
    Set gcolEnderecoEnv = New Collection
    
    For Each vValor In gcolParcelas
        gcolParcelasEnv.Add vValor
    Next
    
    For Each vValor In gcolTitulos
        gcolTitulosEnv.Add vValor
    Next
    
    For Each vValor In gcolClientes
        gcolClientesEnv.Add vValor
    Next
    
    For Each vValor In gcolEndereco
        gcolEnderecoEnv.Add vValor
    Next
    
    For Each vValor In gcolFiliaisCli
        gcolFiliaisCliEnv.Add vValor
    Next
       
    For iLinha = 1 To objGridItens.iLinhasExistentes
                      
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
        
            Set objEnvioEmail = New ClassEnvioDeEmail
            colEnvioEmail.Add objEnvioEmail
        
            objEnvioEmail.sCC = GridItens.TextMatrix(iLinha, iGrid_CC_Col)
            objEnvioEmail.sEmail = GridItens.TextMatrix(iLinha, iGrid_Email_Col)
            objEnvioEmail.sAssunto = GridItens.TextMatrix(iLinha, iGrid_Assunto_Col)
            objEnvioEmail.sMensagem = GridItens.TextMatrix(iLinha, iGrid_Mensagem_Col)
            objEnvioEmail.sAnexo = GridItens.TextMatrix(iLinha, iGrid_Anexo_Col)
            objEnvioEmail.iLinha = iLinha
            
            Set objEnvioEmail.objTela = Me
            
            Set objCobrEmail = gcolModeloLinha.Item(iLinha)
            
            objEnvioEmail.lNumIntDocModelo = objCobrEmail.lNumIntDoc
            
            objEnvioEmail.iConfirmacaoLeitura = objCobrEmail.iConfirmacaoLeitura
            objEnvioEmail.sEmailResp = objCobrEmail.sEmailResp
            
            If giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
                lNumTitulo = gcolTitulos.Item(iLinha).lNumTitulo
                iNumParcela = gcolParcelas.Item(iLinha).iNumParcela
                lCliente = gcolTitulos.Item(iLinha).lFornecedor
                lNumIntParcela = gcolParcelas.Item(iLinha).lNumIntDoc
                iGeraRelac = DESMARCADO
                iFilialCliRelac = gcolTitulos.Item(iLinha).iFilial
                
                objEnvioEmail.iTipoDoc = EMAIL_TIPODOC_PARCELASPAG
                objEnvioEmail.lNumIntDoc = gcolParcelas.Item(iLinha).lNumIntDoc
                
            ElseIf giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
                lNumTitulo = gcolTitulos.Item(iLinha).lNumNotaFiscal
                iNumParcela = 0
                lCliente = gcolTitulos.Item(iLinha).lFornecedor
                lNumIntParcela = 0
                iGeraRelac = DESMARCADO
                iFilialCliRelac = gcolTitulos.Item(iLinha).iFilial
                
                objEnvioEmail.iTipoDoc = EMAIL_TIPODOC_NFSPAG
                objEnvioEmail.lNumIntDoc = gcolTitulos.Item(iLinha).lNumIntDoc
                
            ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
                lNumTitulo = gcolTitulos.Item(iLinha).lNumNotaFiscal
                iNumParcela = 0
                lNumIntParcela = 0
                If gcolTitulos.Item(iLinha).lCliente <> 0 Then
                    iGeraRelac = MARCADO
                    lCliente = gcolTitulos.Item(iLinha).lCliente
                    iFilialCliRelac = gcolTitulos.Item(iLinha).iFilialCli
                Else
                    iGeraRelac = DESMARCADO
                    lCliente = gcolTitulos.Item(iLinha).lFornecedor
                    iFilialCliRelac = gcolTitulos.Item(iLinha).iFilialForn
                End If
                
                objEnvioEmail.iTipoDoc = EMAIL_TIPODOC_NF
                objEnvioEmail.lNumIntDoc = gcolTitulos.Item(iLinha).lNumIntDoc
                
            Else
                lNumTitulo = gcolTitulos.Item(iLinha).lNumTitulo
                iNumParcela = gcolParcelas.Item(iLinha).iNumParcela
                lCliente = gcolTitulos.Item(iLinha).lCliente
                lNumIntParcela = gcolParcelas.Item(iLinha).lNumIntDoc
                iGeraRelac = MARCADO
                iFilialCliRelac = gcolTitulos.Item(iLinha).iFilial
                
                objEnvioEmail.iTipoDoc = EMAIL_TIPODOC_PARCELASREC
                objEnvioEmail.lNumIntDoc = gcolParcelas.Item(iLinha).lNumIntDoc
            End If
              
            
            'Preenche os dados adicionais de cobrança
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
                
                objEnvioEmail.iCobranca = MARCADO
              
                Set objFilialEmpresa = New AdmFiliais
                
                objFilialEmpresa.iCodFilial = gcolTitulosEnv.Item(iLinha).iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196974
                
                objEnvioEmail.colDadosAdicionais.Add "CEDENTE", "", 255, "CEDENTE"
                objEnvioEmail.colDadosAdicionais.Add "NOMECEDENTE", objFilialEmpresa.sNome, 255, "NOMECEDENTE"
                objEnvioEmail.colDadosAdicionais.Add "IDCEDENTE", "", 255, "IDCEDENTE"
                objEnvioEmail.colDadosAdicionais.Add "NUMDOC", CStr(lNumTitulo), 255, "NUMDOC"
                objEnvioEmail.colDadosAdicionais.Add "NOMESACADO", gcolClientes.Item(iLinha).sRazaoSocial, 255, "NOMESACADO"
                objEnvioEmail.colDadosAdicionais.Add "CGCSACADO", gcolClientes.Item(iLinha).sCgc, 255, "CGCSACADO"
                objEnvioEmail.colDadosAdicionais.Add "ENDERECOSACADO", gcolEndereco.Item(iLinha).sEndereco, 255, "ENDERECOSACADO"
                objEnvioEmail.colDadosAdicionais.Add "BAIRROSACADO", gcolEndereco.Item(iLinha).sBairro, 255, "BAIRROSACADO"
                objEnvioEmail.colDadosAdicionais.Add "CIDADESACADO", gcolEndereco.Item(iLinha).sCidade, 255, "CIDADESACADO"
                objEnvioEmail.colDadosAdicionais.Add "UFSACADO", gcolEndereco.Item(iLinha).sSiglaEstado, 255, "UFSACADO"
                objEnvioEmail.colDadosAdicionais.Add "CEPSACADO", gcolEndereco.Item(iLinha).sCEP, 255, "CEPSACADO"
                objEnvioEmail.colDadosAdicionais.Add "DATAVENC", Format(gcolParcelas.Item(iLinha).dtDataVencimento, "dd/mm/yyyy"), 255, "DATAVENC"
                objEnvioEmail.colDadosAdicionais.Add "VALOR", Format(gcolParcelas.Item(iLinha).dSaldo, "STANDARD"), 255, "VALOR"
                objEnvioEmail.colDadosAdicionais.Add "JUROS", "", 255, "JUROS"
                objEnvioEmail.colDadosAdicionais.Add "MULTA", "", 255, "MULTA"
                objEnvioEmail.colDadosAdicionais.Add "DESCONTO", "", 255, "DESCONTO"
                objEnvioEmail.colDadosAdicionais.Add "OUTROSACRESCIMOS", "", 255, "OUTROSACRESCIMOS"
                objEnvioEmail.colDadosAdicionais.Add "DEMONSTRATIVO", "", 255, "DEMONSTRATIVO"
                objEnvioEmail.colDadosAdicionais.Add "INSTRUCOESCAIXA", "Após o vencimento cobrar multa de 1% a.m pró rata die + juros de mora de 0,33% ao dia. " & vbNewLine & "Não receber após 10 dias de atraso.", 255, "INSTRUCOESCAIXA"
                objEnvioEmail.colDadosAdicionais.Add "NOSSONUMERO", CStr(lNumTitulo), 255, "NOSSONUMERO"
                
            End If
            
            lErro = CF2(Me, "CE_BotaoEmail_Click", iGeraRelac)
            If lErro <> SUCESSO Then gError 196974
            
            objEnvioEmail.lClienteRelac = lCliente
            objEnvioEmail.lNumIntDocParc = lNumIntParcela
            objEnvioEmail.iGeraRelac = iGeraRelac
            objEnvioEmail.iFilialCliRelac = iFilialCliRelac
            objEnvioEmail.sTextoRelac = objEnvioEmail.sMensagem
            objEnvioEmail.sModelo = objCobrEmail.sModelo
            objEnvioEmail.sDe = objCobrEmail.sDe
            objEnvioEmail.sNomeExibicao = objCobrEmail.sNomeExibicao
            
            sMsgAux = ""
            If InStr(1, objEnvioEmail.sTextoRelac, "<Mensagem automática confome modelo") <> 0 Then
            
                If objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO Then
                    sMsgAux = "Email de aviso da fatura "
                ElseIf objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA Then
                    sMsgAux = "Email de cobrança para fatura "
                ElseIf objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO Then
                    sMsgAux = "Email de agradecimento do pagto da fatura "
                ElseIf objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA Then
                    sMsgAux = "Email de cobrança de envio da Fatura "
                ElseIf objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP Then
                    sMsgAux = "Email de aviso de pagamento da fatura "
                ElseIf objCobrEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_NFE Then
                    sMsgAux = "Email de envio da Nota Fiscal Eletrônica "
                End If

                If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                    sMsgAux = sMsgAux & CStr(lNumTitulo) & "\" & CStr(iNumParcela)
                Else
                    sMsgAux = sMsgAux & CStr(lNumTitulo)
                End If
                sMsgAux = sMsgAux & ", confome modelo '" & objCobrEmail.sModelo & "', enviado para o email '" & objEnvioEmail.sEmail & "', com cópia para '" & objEnvioEmail.sCC & "', assunto '" & objEnvioEmail.sAssunto & "' e contendo como anexo o arquivo '" & objEnvioEmail.sAnexo & "'."
            
                objEnvioEmail.sTextoRelac = sMsgAux
            
            End If
        
        End If
       
    Next
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 196974

    lErro = CF("Rotina_Envia_Emails_Batch", sNomeArqParam, colEnvioEmail)
    If lErro <> SUCESSO Then gError 196975
    
    'GL_objMDIForm.MousePointer = vbDefault

    'Call Rotina_Aviso(vbOKOnly, "AVISO_EMAIL_ENVIADO_COM_SUCESSO")
    
    iAlterado = 0

    Exit Sub
    
Erro_BotaoEmail_Click:

    'GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 187038
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_PREENCHIDO_GRID", gErr, iLinha)
        
        Case 187039, 189317, 189318, 189319, 189413, 189416, 189418, 189420
            
        Case 187118
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_SELECIONADA", gErr)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187040)

    End Select
    
    Exit Sub
    
End Sub

Function CE_BotaoEmail_Click(ByVal objCT As Object, iGeraRelac As Integer)
    CE_BotaoEmail_Click = SUCESSO
End Function

Public Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
            lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        Else
            lErro = TP_Fornecedor_Le2(ClienteFinal, objForn, 0)
        End If
        If lErro <> SUCESSO Then gError 187041

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 187041
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187042)

    End Select

End Sub

Public Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
            lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        Else
            lErro = TP_Fornecedor_Le2(ClienteInicial, objForn, 0)
        End If
        If lErro <> SUCESSO Then gError 187043

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 187043
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187044)

    End Select

End Sub

Public Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor

On Error GoTo Erro_LabelClienteAte_Click
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
        objForn.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
        Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)
    Else
        Call Chama_Tela("FornecedoresLista", colSelecao, objForn, objEventoClienteFim)
    End If
    
   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 187045)

    End Select

    Exit Sub

End Sub

Public Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor

On Error GoTo Erro_LabelClienteDe_Click
    
    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    If giTipoTela = TIPOTELA_EMAIL_AGRADECIMENTO Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_NFE Then
        Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)
    Else
        Call Chama_Tela("FornecedoresLista", colSelecao, objForn, objEventoClienteInic)
    End If
    
   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187046)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

    ClienteFinal.Text = CStr(obj1.lCodigo)
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)
    
    ClienteInicial.Text = CStr(obj1.lCodigo)
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub SaldoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_SaldoAte_Validate

    'Se a data SaldoAte está preenchida
    If Len(Trim(SaldoAte.Text)) > 0 Then

        'Verifica se a data SaldoAte é válida
        lErro = Valor_Positivo_Critica_Double(SaldoAte.Text)
        If lErro <> SUCESSO Then gError 187047

    End If

    Exit Sub

Erro_SaldoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 187047

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187048)

    End Select

    Exit Sub

End Sub

Public Sub SaldoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_SaldoDe_Validate

    'Se a data SaldoDe está preenchida
    If Len(Trim(SaldoDe.Text)) > 0 Then

        'Verifica se a data SaldoAte é válida
        lErro = Valor_Positivo_Critica_Double(SaldoDe.Text)
        If lErro <> SUCESSO Then gError 187049

    End If

    Exit Sub

Erro_SaldoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 187049

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187050)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    If giTipoTela <> TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        'Le os Tipos de Documentos utilizados em Titulos a Receber
        lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    Else
        lErro = CF("TiposDocumento_Le_TituloPag", colTipoDocumento)
    End If
    If lErro <> SUCESSO Then gError 187051
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 187051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187052)

    End Select

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional sValorTexto As String, Optional ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sMes As String
Dim objFilialEmpresa As New AdmFiliais
Dim colNF As New ColNFsPag
Dim objNFPag As ClassNFsPag
Dim objNF As ClassNFiscal, sTela As String
Dim sTextoAux As String
Dim iIndice As Integer
Dim gcolParcelasMen As New Collection
Dim gcolTitulosMen As New Collection
Dim gcolClientesMen As New Collection
Dim gcolFiliaisCliMen As New Collection
Dim vValor As Variant
Dim sRetorno As String
Dim lRetorno As Long
Dim iPos As Integer
Dim sBoleto As String
Dim objNFiscal As ClassNFiscal
Dim objTipoDocInfo As ClassTipoDocInfo
Dim objTransp As ClassTransportadora
Dim objVend As ClassVendedor
Dim objEndereco As ClassEndereco
Dim sEmail As String, lSeqEvento As Long
Dim sMotivo As String, sCodVerificacao As String, lNumNFe As Long
Dim sLink As String, objNFSeLinks As ClassNFSeLinks
Dim colSeqEvento As Collection

On Error GoTo Erro_Calcula_Mnemonico
   
    'Se não passou a linha o calcula mnemonico vem da tela logo segue a ordenação da tela
    If iLinha = 0 Then
        For Each vValor In gcolParcelas
            gcolParcelasMen.Add vValor
        Next
        
        For Each vValor In gcolTitulos
            gcolTitulosMen.Add vValor
        Next
        
        For Each vValor In gcolClientes
            gcolClientesMen.Add vValor
        Next
        
        For Each vValor In gcolFiliaisCli
            gcolFiliaisCliMen.Add vValor
        Next
    Else ' Se passou a linha vem do batch logo deve seguir a ordenação do momento do envio
        For Each vValor In gcolParcelasEnv
            gcolParcelasMen.Add vValor
        Next
        
        For Each vValor In gcolTitulosEnv
            gcolTitulosMen.Add vValor
        Next
        
        For Each vValor In gcolClientesEnv
            gcolClientesMen.Add vValor
        Next
        
        For Each vValor In gcolFiliaisCliEnv
            gcolFiliaisCliMen.Add vValor
        Next
    End If

    objFilialEmpresa.iCodFilial = gcolTitulosMen.Item(giLinhaAtual).iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 196985

    sValorTexto = ""
    
    If iLinha <> 0 Then giLinhaAtual = iLinha

    Select Case UCase(objMnemonicoValor.sMnemonico)
                                                      
        Case UCase(DATA_VENCIMENTO)
            If giTipoTela = TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add DATA_NULA
                sValorTexto = "__/__/____"
            Else
                objMnemonicoValor.colValor.Add gcolParcelasMen.Item(giLinhaAtual).dtDataVencimento
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dtDataVencimento, "dd/mm/yyyy")
            End If
        Case UCase(DATA_VENCIMENTO_REAL)
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Then
                objMnemonicoValor.colValor.Add gcolParcelasMen.Item(giLinhaAtual).dtDataVencimento
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dtDataVencimento, "dd/mm/yyyy")
            ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add DATA_NULA
                sValorTexto = "__/__/____"
            Else
                objMnemonicoValor.colValor.Add gcolParcelasMen.Item(giLinhaAtual).dtDataVencimentoReal
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dtDataVencimentoReal, "dd/mm/yyyy")
            End If

        Case UCase(NUMERO_PARCELA)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add gcolParcelasMen.Item(giLinhaAtual).iNumParcela
                sValorTexto = CStr(gcolParcelasMen.Item(giLinhaAtual).iNumParcela)
            Else
                objMnemonicoValor.colValor.Add 1
                sValorTexto = "1"
            End If
            
        Case UCase(BOLETO_CORPORATOR)
        
            If giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_COBRANCA Then
            
                lErro = CF("Boleto_Gera_Arquivo", sBoleto, gcolTitulosMen.Item(giLinhaAtual), gcolParcelasMen.Item(giLinhaAtual))
                If lErro <> SUCESSO Then gError 196985
                
                objMnemonicoValor.colValor.Add sBoleto
                sValorTexto = sBoleto
            End If
            
        Case UCase(NUMERO_TITULO)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).lNumTitulo
                sValorTexto = CStr(gcolTitulosMen.Item(giLinhaAtual).lNumTitulo)
            Else
                objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal
                sValorTexto = CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)
            End If
            
        Case UCase(NUMERO_TITULO_TRV)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                sValorTexto = String(6 - Len(CStr(gcolTitulosMen.Item(giLinhaAtual).lNumTitulo)), "0") & CStr(gcolTitulosMen.Item(giLinhaAtual).lNumTitulo)
                objMnemonicoValor.colValor.Add sValorTexto
            Else
                sValorTexto = String(6 - Len(CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)), "0") & CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)
                objMnemonicoValor.colValor.Add sValorTexto
            End If
            
        Case UCase(NUMERO_TITULO_TRP)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                sValorTexto = String(6 - Len(CStr(gcolTitulosMen.Item(giLinhaAtual).lNumTitulo)), "0") & CStr(gcolTitulosMen.Item(giLinhaAtual).lNumTitulo)
                objMnemonicoValor.colValor.Add sValorTexto
            Else
                sValorTexto = String(6 - Len(CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)), "0") & CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)
                objMnemonicoValor.colValor.Add sValorTexto
            End If

        Case UCase(RAZAO_CLIENTE)
            objMnemonicoValor.colValor.Add gcolClientesMen.Item(giLinhaAtual).sRazaoSocial
            sValorTexto = gcolClientesMen.Item(giLinhaAtual).sRazaoSocial
            
        Case UCase(CNPJ_CLIENTE)
            objMnemonicoValor.colValor.Add Format(gcolFiliaisCliMen.Item(giLinhaAtual).sCgc, "00\.000\.000\/0000-00; ; ; ")
            sValorTexto = Format(gcolFiliaisCliMen.Item(giLinhaAtual).sCgc, "00\.000\.000\/0000-00; ; ; ")

        Case UCase(SALDO_PARCELA)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add Format(gcolParcelasMen.Item(giLinhaAtual).dSaldo, "STANDARD")
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dSaldo, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(gcolParcelasMen.Item(giLinhaAtual).dValorTotal, "STANDARD")
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dValorTotal, "STANDARD")
            End If
            
        Case UCase(VALOR_PARCELA), UCase(VALOR_TOTAL)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add Format(gcolParcelasMen.Item(giLinhaAtual).dValor, "STANDARD")
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dValor, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(gcolParcelasMen.Item(giLinhaAtual).dValorTotal, "STANDARD")
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dValorTotal, "STANDARD")
            End If
            
        Case UCase(DIA_ATUAL)
            objMnemonicoValor.colValor.Add Day(gdtDataAtual)
            sValorTexto = CStr(Day(gdtDataAtual))
        
        Case UCase(MES_NOME)
            Call MesNome(Month(gdtDataAtual), sMes)
            objMnemonicoValor.colValor.Add sMes
            sValorTexto = sMes
        
        Case UCase(MES_ATUAL)
            objMnemonicoValor.colValor.Add Month(gdtDataAtual)
            sValorTexto = CStr(Month(gdtDataAtual))
        
        Case UCase(ANO_ATUAL)
            objMnemonicoValor.colValor.Add Year(gdtDataAtual)
            sValorTexto = CStr(Year(gdtDataAtual))
        
        Case UCase(DATA_ATUAL)
            objMnemonicoValor.colValor.Add gdtDataAtual
            sValorTexto = Format(gdtDataAtual, "dd/mm/yyyy")
            
        Case UCase(DATA_BAIXA)
            If giTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add gcolParcelasMen.Item(giLinhaAtual).dtDataUltimaBaixa
                sValorTexto = Format(gcolParcelasMen.Item(giLinhaAtual).dtDataUltimaBaixa, "dd/mm/yyyy")
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
                sValorTexto = "__/__/____"
            End If
            
        Case UCase(DATA_EMISSAO)
            objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).dtDataEmissao
            sValorTexto = Format(gcolTitulosMen.Item(giLinhaAtual).dtDataEmissao, "dd/mm/yyyy")
        
        Case UCase(NFE_CHAVE), UCase(NFSE_LINK_CONSULTA), UCase(NFSE_LINK_SITE), UCase(NFSE_LINK_VERIFICACAO), UCase(NFE_CHAVE2), UCase(NFSE_LINKS_PADRAO)
        
            Set objNF = New ClassNFiscal
            sLink = ""
            sCodVerificacao = ""
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then
                
                lErro = CF("Titulo_Le_DocumentoOriginal", gcolTitulosMen.Item(giLinhaAtual).lNumIntDoc, CPR_TITULO_RECEBER, objNF, sTela)
                If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
                
                If lErro = SUCESSO Then
                    lErro = CF("NFiscal_Le", objNF)
                    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                    
                    sCodVerificacao = objNF.sCodVerificacaoNFe
                End If
                
            ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
                sCodVerificacao = gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe
            End If
            
            If UCase(objMnemonicoValor.sMnemonico) = UCase(NFE_CHAVE) Then
                sValorTexto = sCodVerificacao
            ElseIf UCase(objMnemonicoValor.sMnemonico) = UCase(NFE_CHAVE2) Then
                sValorTexto = Replace(Replace(Replace(Replace(Replace(sCodVerificacao, "-", ""), " ", ""), ".", ""), "/", ""), "\", "")
            Else
            
                If Len(Trim(sCodVerificacao)) > 0 Then
                    Set objNFSeLinks = New ClassNFSeLinks
                    objNFSeLinks.iFilialEmpresa = objNF.iFilialEmpresa
                    lErro = CF("NFSeLinks_Le", objNFSeLinks)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                        If UCase(objMnemonicoValor.sMnemonico) = UCase(NFSE_LINK_CONSULTA) Then
                            sLink = objNFSeLinks.sLinkConsulta & objNFSeLinks.sParamConsulta
                        ElseIf UCase(objMnemonicoValor.sMnemonico) = UCase(NFSE_LINK_SITE) Then
                            sLink = objNFSeLinks.sLinkSite
                        ElseIf UCase(objMnemonicoValor.sMnemonico) = UCase(NFSE_LINK_VERIFICACAO) Then
                            sLink = objNFSeLinks.sLinkVerificacao
                        ElseIf UCase(objMnemonicoValor.sMnemonico) = UCase(NFSE_LINKS_PADRAO) Then
                            sLink = "<br><br>Esse boleto refere-se a nota fiscal de serviço eletrônica que pode ser consultada pelo link abaixo<br><a href=""ZZZNFSe_Link_ConsultaZZZ"">ZZZNFSe_Link_ConsultaZZZ</a><br><br>Ou através do site<br><a href=""ZZZNFSe_LinkVerificacaoZZZ"">ZZZNFSe_LinkVerificacaoZZZ</a><br>Informando:<br>CNPJ do Prestador = ZZZCNPJ_EmpZZZ<br>Número da NFS-e = ZZZNumero_NFEZZZ<br>Código de Verificação = ZZZNFe_ChaveZZZ"
                        End If
                    End If
                
                End If
            
                sValorTexto = sLink
            End If
            objMnemonicoValor.colValor.Add sValorTexto
            
        Case UCase(NUMERO_NFE)
        
            If giTipoTela = TIPOTELA_EMAIL_COBRANCA Or giTipoTela = TIPOTELA_EMAIL_AVISO_COBRANCA Then

                Set objNF = New ClassNFiscal
                
                lErro = CF("Titulo_Le_DocumentoOriginal", gcolTitulosMen.Item(giLinhaAtual).lNumIntDoc, CPR_TITULO_RECEBER, objNF, sTela)
                If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
                
                sValorTexto = ""
                If lErro = SUCESSO Then
                    
                    lErro = CF("NFiscal_Le", objNF)
                    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                    
                    If objNF.iRecibo = DESMARCADO Then
                        sValorTexto = CStr(objNF.lNumNotaFiscal)
                    Else
                        sValorTexto = CStr(objNF.lNumNFe)
                    End If
                    objMnemonicoValor.colValor.Add StrParaLong(sValorTexto)
                    If objNF.iStatus = STATUS_CANCELADO Then
                        sValorTexto = sValorTexto & "(CANCELADA)"
                    End If
                Else
                    objMnemonicoValor.colValor.Add StrParaLong(sValorTexto)
                End If
                
            ElseIf giTipoTela = TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal
                sValorTexto = CStr(gcolTitulosMen.Item(giLinhaAtual).lNumNotaFiscal)
                
                If gcolTitulosMen.Item(giLinhaAtual).iStatus = STATUS_CANCELADO Then
                    sValorTexto = sValorTexto & "(CANCELADA)"
                End If
                lErro = CF("NFe_Le_Cce", gcolTitulosMen.Item(giLinhaAtual).lNumIntDoc, lSeqEvento)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                
                If lErro = SUCESSO Then
                    sValorTexto = sValorTexto & "(c/Cce)"
                End If
            Else
                objMnemonicoValor.colValor.Add 0
                sValorTexto = ""
            End If
            
        Case UCase(ARQ_XML)
            If giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add ""
                sValorTexto = ""
            Else
                If gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "1.10" Then
                    objMnemonicoValor.colValor.Add "NFe" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & ".xml"
                    sValorTexto = "NFe" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & ".xml"
                ElseIf gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "2.00" Then
                    objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-nfe.xml"
                    sValorTexto = gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-nfe.xml"
                Else '3.10\4.00 ou superior
                    objMnemonicoValor.colValor.Add gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-procNfe.xml"
                    sValorTexto = gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-procNfe.xml"
                End If

            End If

        Case UCase(DIR_XML)
            If giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add ""
                sValorTexto = ""
            Else
                lErro = CF("NFe_Obtem_Dir_Xml", sRetorno)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                objMnemonicoValor.colValor.Add sRetorno
                sValorTexto = sRetorno
            
            End If
            
        Case UCase(CAMINHO_XML)

            If giTipoTela <> TIPOTELA_EMAIL_NFE Then
                objMnemonicoValor.colValor.Add ""
                sValorTexto = ""
            Else
                lErro = CF("NFe_Obtem_Dir_Xml", sRetorno)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                If gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "1.10" Then
                    sValorTexto = sRetorno & "NFe" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & ".xml"
                ElseIf gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "2.00" Then
                    sValorTexto = sRetorno & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-nfe.xml"
                Else '3.10\4.00 ou superior
                    sValorTexto = sRetorno & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-procNfe.xml"
                End If
                
                If gcolTitulosMen.Item(giLinhaAtual).iStatus = STATUS_CANCELADO Then
                    If gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "1.10" Then
                        sValorTexto = sValorTexto & ";" & sRetorno & "CancNFe" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & ".xml"
                    ElseIf gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "2.00" Then
                        sValorTexto = sValorTexto & ";" & sRetorno & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-can.xml"
                    Else '3.10\4.00 ou superior
                        sValorTexto = sValorTexto & ";" & sRetorno & "110111-" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & "-1-procEventoNfe.xml"
                    End If
                End If
                
                Set colSeqEvento = New Collection
                
                lErro = CF("NFe_Le_Cce2", gcolTitulosMen.Item(giLinhaAtual).lNumIntDoc, colSeqEvento)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'For Each vValor In colSeqEvento
                If colSeqEvento.Count <> 0 Then
                
                    vValor = colSeqEvento.Item(1)
                
                    If gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "1.10" Or gcolTitulosMen.Item(giLinhaAtual).sStatusNFeFed = "2.00" Then
                        sValorTexto = sValorTexto & ";" & sRetorno & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & SEPARADOR & CStr(vValor) & "-nfe-cce.xml"
                    Else '3.10\4.00 ou superior
                        sValorTexto = sValorTexto & ";" & sRetorno & "110110-" & gcolTitulosMen.Item(giLinhaAtual).sCodVerificacaoNFe & SEPARADOR & CStr(vValor) & "-procEventoNfe.xml"
                    End If
                End If
                'Next
                
                objMnemonicoValor.colValor.Add left(sValorTexto, 250)
                
            End If
            
        Case UCase(EMAIL_TRANSP_SAIDA)
        
            sEmail = ""
        
            If giTipoTela = TIPOTELA_EMAIL_NFE Then
            
                Set objNFiscal = gcolTitulosMen.Item(giLinhaAtual)
                Set objTipoDocInfo = New ClassTipoDocInfo

                objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
                
                If objTipoDocInfo.iTipo = 2 And objNFiscal.iCodTransportadora <> 0 Then
                
                    Set objTransp = New ClassTransportadora
                    Set objEndereco = New ClassEndereco
                
                    objTransp.iCodigo = objNFiscal.iCodTransportadora
                    
                    lErro = CF("Transportadora_Le", objTransp)
                    If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM

                    'Carrega lEndereco em objTransportadora
                    objEndereco.lCodigo = objTransp.lEndereco
                
                    'Lê o endereço a partir do Código
                    lErro = CF("Endereco_Le", objEndereco)
                    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM

                    sEmail = objEndereco.sEmail
                
                End If
                
                If objTipoDocInfo.iTipo = 2 And objNFiscal.iCodTranspRedesp <> 0 And objNFiscal.iCodTranspRedesp <> objNFiscal.iCodTransportadora Then
                
                    Set objTransp = New ClassTransportadora
                    Set objEndereco = New ClassEndereco
                
                    objTransp.iCodigo = objNFiscal.iCodTranspRedesp
                    
                    lErro = CF("Transportadora_Le", objTransp)
                    If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM

                    'Carrega lEndereco em objTransportadora
                    objEndereco.lCodigo = objTransp.lEndereco
                
                    'Lê o endereço a partir do Código
                    lErro = CF("Endereco_Le", objEndereco)
                    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM

                    If Len(Trim(objEndereco.sEmail)) > 0 Then
                        If sEmail <> "" Then sEmail = sEmail & ";"
                        sEmail = sEmail & objEndereco.sEmail
                    End If
                
                End If
            
            End If
            objMnemonicoValor.colValor.Add sEmail
            sValorTexto = sEmail
            
        Case UCase(EMAIL_VENDEDOR)
        
            sEmail = ""
        
            If giTipoTela = TIPOTELA_EMAIL_NFE Then
            
                Set objNFiscal = gcolTitulosMen.Item(giLinhaAtual)
                
                If objNFiscal.lCliente <> 0 Then
            
                    Set objVend = New ClassVendedor
                                
                    objVend.iCodigo = gcolFiliaisCliMen.Item(giLinhaAtual).iVendedor
                    
                    If objVend.iCodigo <> 0 Then
                    
                        lErro = CF("Vendedor_Le", objVend)
                        If lErro <> SUCESSO And lErro <> 12582 Then gError ERRO_SEM_MENSAGEM
                        
                        Set objEndereco = New ClassEndereco
                        
                        objEndereco.lCodigo = objVend.lEndereco
                    
                        'Endereco
                        lErro = CF("Endereco_Le", objEndereco)
                        If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                        
                        sEmail = objEndereco.sEmail
                        
                    End If
                    
                End If
            
            End If
            objMnemonicoValor.colValor.Add sEmail
            sValorTexto = sEmail
            
        Case UCase(NOME_EMP)
            objMnemonicoValor.colValor.Add gsNomeEmpresa
            sValorTexto = gsNomeEmpresa
            
        Case UCase(CNPJ_EMP)
            objMnemonicoValor.colValor.Add Format(objFilialEmpresa.sCgc, "00\.000\.000\/0000-00; ; ; ")
            sValorTexto = Format(objFilialEmpresa.sCgc, "00\.000\.000\/0000-00; ; ; ")
            
        Case UCase(IM_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.sInscricaoMunicipal
            sValorTexto = objFilialEmpresa.sInscricaoMunicipal
            
        Case UCase(IM_EMP2)
            objMnemonicoValor.colValor.Add Replace(Replace(Replace(Replace(Replace(objFilialEmpresa.sInscricaoMunicipal, "-", ""), " ", ""), ".", ""), "/", ""), "\", "")
            sValorTexto = Replace(Replace(Replace(Replace(Replace(objFilialEmpresa.sInscricaoMunicipal, "-", ""), " ", ""), ".", ""), "/", ""), "\", "")
            
        Case UCase(ENDERECO_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.objEndereco.sEndereco
            sValorTexto = objFilialEmpresa.objEndereco.sEndereco
            
        Case UCase(BAIRRO_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.objEndereco.sBairro
            sValorTexto = objFilialEmpresa.objEndereco.sBairro
            
        Case UCase(CIDADE_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.objEndereco.sCidade
            sValorTexto = objFilialEmpresa.objEndereco.sCidade
            
        Case UCase(UF_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.objEndereco.sSiglaEstado
            sValorTexto = objFilialEmpresa.objEndereco.sSiglaEstado
            
        Case UCase(CEP_EMP)
            objMnemonicoValor.colValor.Add objFilialEmpresa.objEndereco.sCEP
            sValorTexto = objFilialEmpresa.objEndereco.sCEP
            
        Case UCase(LISTA_NFSPAG)
        
            sTextoAux = ""
        
            If giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
                
                lErro = CF("NfsPag_Le_FaturaPagar2", gcolTitulosMen.Item(giLinhaAtual), colNF)
                If lErro <> SUCESSO And lErro <> 26020 Then gError 196985
                
                If colNF.Count > 0 Then
                    
                    iIndice = 0
                    For Each objNFPag In colNF
                        iIndice = iIndice + 1
                        
                        If iIndice = 1 Then
                            sTextoAux = CStr(objNFPag.lNumNotaFiscal)
                        ElseIf iIndice = colNF.Count Then
                            sTextoAux = sTextoAux & " e " & CStr(objNFPag.lNumNotaFiscal)
                        Else
                            sTextoAux = sTextoAux & ", " & CStr(objNFPag.lNumNotaFiscal)
                        End If
                    
                    Next
                
                End If
            
            End If
            
            objMnemonicoValor.colValor.Add sTextoAux
            sValorTexto = sTextoAux
            
        Case Else
            lErro = CF("CE_Calcula_Mnemonico_Cust", objMnemonicoValor, giTipoTela, gcolTitulosMen.Item(giLinhaAtual), sValorTexto)
            If lErro <> SUCESSO Then gError 187163

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case 424
            Call Rotina_Erro(vbOKOnly, "ERRO_ENVIO_DE_EMAIL_TELA_FECHADA", gErr)
    
        Case 196985

        Case 187163
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_NAO_ENCONTRADO", gErr, objMnemonicoValor.sMnemonico)

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187164)

    End Select

    Exit Function

End Function

Public Function CE_Calcula_Mnemonico_Cust(objMnemonicoValor As ClassMnemonicoValor, ByVal iTipoTela As Integer, ByVal objDoc As Object, sValorTexto As String) As Long
    CE_Calcula_Mnemonico_Cust = SUCESSO
End Function

'??? só para manter a compatibilidade binaria
Public Function Verifica_Existencia_Arquivo(ByVal sArquivo As String) As Long
    Verifica_Existencia_Arquivo = Verifica_Existencia_Arquivo2(sArquivo)
End Function

Private Function Verifica_Existencia_Arquivo2(ByVal sArquivo As String, Optional ByVal bSemMsgErro As Boolean = False) As Long

Dim lErro As Long
Dim iPos As Integer
Dim iPosAnt As Integer
Dim sArqAux As String

On Error GoTo Erro_Verifica_Existencia_Arquivo2

    If Len(Trim(sArquivo)) > 0 Then

        sArqAux = sArquivo
        iPosAnt = 0
        iPos = InStr(iPosAnt + 1, sArquivo, ";")
        Do While iPos <> 0
            sArqAux = Mid(sArquivo, iPosAnt + 1, iPos - (iPosAnt + 1))
            Open Trim(sArqAux) For Input As #1
            Close #1
            iPosAnt = iPos
            iPos = InStr(iPosAnt + 1, sArquivo, ";")
        Loop
        If iPosAnt <> 0 Then
            sArqAux = Mid(sArquivo, iPosAnt + 1)
        End If
        Open Trim(sArqAux) For Input As #1
        Close #1
        
    End If

    Verifica_Existencia_Arquivo2 = SUCESSO

    Exit Function

Erro_Verifica_Existencia_Arquivo2:

    Verifica_Existencia_Arquivo2 = gErr

    Select Case gErr
    
        Case 53, 52
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, sArqAux)

        Case Else
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189322)

    End Select

    Exit Function

End Function

Public Sub AtrasoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtrasoDe_Validate

    If Len(Trim(AtrasoDe.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(AtrasoDe.Text)
    If lErro <> SUCESSO Then gError 189402
   
    Exit Sub

Erro_AtrasoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 189402
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189403)

    End Select

    Exit Sub

End Sub

Public Sub AtrasoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtrasoAte_Validate

    If Len(Trim(AtrasoAte.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(AtrasoAte.Text)
    If lErro <> SUCESSO Then gError 189402
   
    Exit Sub

Erro_AtrasoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 189402
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189403)

    End Select

    Exit Sub

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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        End If

    End If

End Sub

Public Sub DataBaixaDe_Change()
    'Registra que houve alteração
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub DataBaixaDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBaixaDe, iAlteradoFiltro)
End Sub

Public Sub DataBaixaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixaDe_Validate

    'Se a data EmissaoFim está preenchida
    If Len(DataBaixaDe.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(DataBaixaDe.Text)
        If lErro <> SUCESSO Then gError 189434
    
    End If

    Exit Sub

Erro_DataBaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 189434

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189435)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 189432

    Exit Sub

Erro_UpDownBaixaDe_DownClick:

    Select Case gErr

        Case 189432 'Erro tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189433)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 189430

    Exit Sub

Erro_UpDownBaixaDe_UpClick:

    Select Case gErr

        Case 189430 'Erro tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189431)

    End Select

    Exit Sub

End Sub

Public Sub DataBaixaAte_Change()
    'Registra que houve alteração
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub DataBaixaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBaixaAte, iAlteradoFiltro)
End Sub

Public Sub DataBaixaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixaAte_Validate

    'Se a data EmissaoFim está preenchida
    If Len(DataBaixaAte.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(DataBaixaAte.Text)
        If lErro <> SUCESSO Then gError 189434
    
    End If

    Exit Sub

Erro_DataBaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 189434

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189435)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 189432

    Exit Sub

Erro_UpDownBaixaAte_DownClick:

    Select Case gErr

        Case 189432 'Erro tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189433)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 189430

    Exit Sub

Erro_UpDownBaixaAte_UpClick:

    Select Case gErr

        Case 189430 'Erro tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189431)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaCliente_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Click

    If Len(Trim(CategoriaCliente.Text)) > 0 Then
        CategoriaClienteDe.Enabled = True
        CategoriaClienteAte.Enabled = True
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteDe)
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteAte)
    Else
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
    End If

    iAlteradoFiltro = REGISTRO_ALTERADO

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168906)

    End Select

    Exit Sub

End Sub

Public Sub Carrega_ComboCategoriaCliente(ByVal objCombo As Object)

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 131995

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        objCombo.AddItem objCategoriaCliente.sCategoria

    Next
    
    Exit Sub

Erro_Carrega_ComboCategoriaCliente:

    Select Case gErr
    
        Case 131995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168907)

    End Select

    Exit Sub

End Sub

Public Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As Object, ByVal objComboItens As Object)

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_Carrega_ComboCategoriaItens

    'Verifica se a CategoriaCliente foi preenchida
    If objComboCategoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = objComboCategoria.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 131994

        objComboItens.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        objComboItens.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            objComboItens.AddItem objCategoriaClienteItem.sItem

        Next
        
        CategoriaClienteTodas.Value = vbFalse
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        objComboItens.ListIndex = -1
        objComboItens.Enabled = False
    
    End If
    
    Exit Sub

Erro_Carrega_ComboCategoriaItens:

    Select Case gErr
    
        Case 131993

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168908)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131998
        
        If lErro <> SUCESSO Then gError 131999
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        CategoriaClienteDe.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Enabled = False
        CategoriaClienteAte.Clear
    End If
    
    iAlteradoFiltro = REGISTRO_ALTERADO
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 131998
         
        Case 131999
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168909)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaClienteItem_Validate(Cancel As Boolean, objCombo As Object)

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Validate

    If Len(objCombo.Text) <> 0 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(objCombo)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131996
        
        If lErro <> SUCESSO Then gError 131997
    
    End If
    
    iAlteradoFiltro = REGISTRO_ALTERADO

    Exit Sub

Erro_CategoriaClienteItem_Validate:

    Cancel = True

    Select Case gErr

        Case 131996
        
        Case 131997
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCombo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168910)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaClienteTodas_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteTodas_Click

    If CategoriaClienteTodas.Value = vbChecked Then
        'Desabilita o combotipo
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Clear
    Else
        CategoriaCliente.Enabled = True
    End If

    Call CategoriaCliente_Click
    
    iAlteradoFiltro = REGISTRO_ALTERADO

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168911)

    End Select

    Exit Sub

End Sub

Public Sub EmailValido_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub IgnoraJaEnviados_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_ComboCobrador_Validate
    
    'Verifica se algum codigo está selecionado
    'If comboCobrador.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(comboCobrador.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = comboCobrador.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 182172
        
        If lErro <> SUCESSO Then gError 182173
        
    End If
    
    Exit Sub
    
Erro_ComboCobrador_Validate:

    Cancel = True

    Select Case gErr
            
        Case 182172
        
        Case 182173 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 182135

    For Each objUsuarios In colUsuarios
        comboCobrador.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 182135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182136)

    End Select

    Exit Function

End Function

Public Sub comboCobradorBordero_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub comboCobradorBordero_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Private Sub objEventoBorderoDe_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca
    
    Set objBorderoCobranca = obj1
    
    If BorderoCobrDe.Enabled = True Then
        BorderoCobrDe.PromptInclude = False
        BorderoCobrDe.Text = objBorderoCobranca.lNumBordero
        BorderoCobrDe.PromptInclude = True
        Call BorderoCobrDe_Validate(bSGECancelDummy)
    End If

    Me.Show

End Sub

Public Sub LabelBorderoCobrDe_Click()

Dim lErro As Long
Dim objBordero As New ClassBorderoCobranca
Dim colSelecao As New Collection

On Error GoTo Erro_LabelBorderoCobrDe_Click

    'Verifica se o cobrador foi preenchido
    If Len(Trim(ComboCobradorBordero.Text)) = 0 Then gError 61308

    If Len(Trim(BorderoCobrDe.Text)) > 0 Then objBordero.lNumBordero = CLng(BorderoCobrDe.Text)
    
    colSelecao.Add Codigo_Extrai(ComboCobradorBordero.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBordero, objEventoBorderoDe)

    Exit Sub

Erro_LabelBorderoCobrDe_Click:

    Select Case Err

        Case 61308
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159412)

    End Select

    Exit Sub

End Sub

Public Sub ComboCobradorBordero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCobrador As New ClassCobrador
Dim iCodigo As Integer

On Error GoTo Erro_ComboCobradorBordero_Validate

    'Verifica se foi preenchida a ComboBox ComboCobrador
    If Len(Trim(ComboCobradorBordero.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ComboCobrador
    If ComboCobradorBordero.Text = ComboCobradorBordero.List(ComboCobradorBordero.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboCobradorBordero, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 61361
    
    If iCodigo = COBRADOR_PROPRIA_EMPRESA Then gError 61366
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 61362

        If lErro <> SUCESSO Then gError 61363 'Não encontrou Cobrador no BD

        'Encontrou Cobrador no BD, coloca no Text da Combo
        ComboCobradorBordero.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 61364

    Exit Sub

Erro_ComboCobradorBordero_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 61361, 61362

    Case 61363  'Não encontrou Cobrador no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")
        If vbMsgRes = vbYes Then
            Call Chama_Tela("Cobradores", objCobrador)
        End If

    Case 61364
        Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, ComboCobradorBordero.Text)

    Case 61366
        Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_PROPRIA_EMPRESA", gErr, iCodigo)
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159417)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Combo_CobradorBordero() As Long
'Carrega a Combo de Cobradores com todos os cobradores menos Propria Empresa

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim ColCobrador As New Collection

On Error GoTo Erro_Carrega_Combo_CobradorBordero

    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 61310

    'Preenche a ComboBox Cobrador com os objetos da coleção de Cobradores
    For Each objCobrador In ColCobrador

        ''If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iCodBanco <> 0 Then
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            ComboCobradorBordero.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            ComboCobradorBordero.ItemData(ComboCobradorBordero.NewIndex) = objCobrador.iCodigo
        End If

    Next

    Carrega_Combo_CobradorBordero = SUCESSO

    Exit Function

Erro_Carrega_Combo_CobradorBordero:

    Carrega_Combo_CobradorBordero = gErr

    Select Case gErr

        Case 61310 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159418)

    End Select

    Exit Function

End Function

Public Sub BorderoCobrDe_Change()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub BorderoCobrDe_Validate(Cancel As Boolean)

Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_BorderoCobrDe_Validate

    'Se o Bordero de Não está Preenchido --> Sai da Sub
    If Len(Trim(BorderoCobrDe.Text)) = 0 Then Exit Sub
    
    'Verifica se o cobrador foi preenchido
    If Len(Trim(ComboCobradorBordero.Text)) = 0 Then gError 61450
    
    objBorderoCobranca.lNumBordero = CLng(BorderoCobrDe.Text)
    objBorderoCobranca.iCobrador = Codigo_Extrai(ComboCobradorBordero.Text)
    
    'Lê o Bordero de Cobranca
    lErro = CF("BorderoCobranca_Le_Cobrador", objBorderoCobranca)
    If lErro <> SUCESSO And lErro <> 61447 Then gError 61440
    
    'Não Encontrou o bordero
    If lErro = 61447 Then gError 61441
    
    objCarteiraCobrador.iCobrador = objBorderoCobranca.iCobrador
    objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
    lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 23551 Then gError 130030
    If lErro <> SUCESSO Then gError 130031
    
    Exit Sub
    
Erro_BorderoCobrDe_Validate:
        
    Cancel = True
    
    Select Case gErr
        
        Case 61357
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERODE_MAIOR_BORDEROATE", gErr)
        
        Case 61440, 130030
        
        Case 130031
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTCOBR_BORDERO", gErr)
        
        Case 61441
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO_COBRADOR", gErr, objBorderoCobranca.lNumBordero, objBorderoCobranca.iCobrador)
        
        Case 61450
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159406)

    End Select
        
    Exit Sub
    
End Sub

Public Sub BorderoCobrDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(BorderoCobrDe)
End Sub

Public Sub cCliente_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub cFixo_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
    If cFixo.Value = vbChecked Then
        emailFixo.Enabled = True
    Else
        emailFixo.Enabled = False
        emailFixo.Text = ""
    End If
End Sub

Public Sub cTransportadora_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Sub optIncluiDanfe_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub
