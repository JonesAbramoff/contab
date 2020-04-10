VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTContatoCliPorEmail"
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

Public sMsg As String

Public iAtualizaTela As Integer

Public sOpcaoAnt As String

Public iAlterado As Integer
Public iAlteradoFiltro As Integer
Public iFrameAtual As Integer

Public gobjInfoUsu As Object

Public iLinhaAnt As Integer

Dim objGridClientes As AdmGrid
Public iGrid_Selecionado_Col As Integer
Public iGrid_Cliente_Col As Integer
Public iGrid_Filial_Col As Integer
Public iGrid_Carta_Col As Integer
Public iGrid_CC_Col As Integer
Public iGrid_Assunto_Col As Integer
Public iGrid_Email_Col As Integer
Public iGrid_Mensagem_Col As Integer
Public iGrid_Anexo_Col As Integer
Public iGrid_ValorCompra_Col As Integer
Public iGrid_DataUCompra_Col As Integer
Public iGrid_DataPCompra_Col As Integer
Public iGrid_DataUContato_Col As Integer

Public gobjContatoCliAnt As ClassContatoCliSel

Public gcolClientes As Collection
Public gcolFiliais As Collection
Public gcolEnderecos As Collection
Public gcolEstatisticas As Collection
Public gcolModelos As Collection
Public giLinhaAtual As Integer

Public gcolClientesEnv As Collection
Public gcolFiliaisEnv As Collection
Public gcolEnderecosEnv As Collection
Public gcolEstatisticasEnv As Collection

Const TAB_SELECAO = 1
Const TAB_CLIENTE = 2

Const FRAMED_FAIXA = 0
Const FRAMED_APENAS = 1
Const FRAMED_ENTRE = 2

Public gbTrazendoDados As Boolean

'Mnemonicos
'Const DATA_VENCIMENTO = "Data_Vencimento"
'Const DATA_VENCIMENTO_REAL = "Data_Vencimento_Real"
'Const NUMERO_PARCELA = "Numero_Parcela"
'Const NUMERO_TITULO = "Numero_Titulo"
Const RAZAO_CLIENTE = "Razao_Cliente"
Const CONTATO_CLIENTE = "Contato_Cliente"
'Const SALDO_PARCELA = "Saldo_Parcela"
'Const VALOR_PARCELA = "Valor_Parcela"
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
'Const LISTA_NFSPAG = "Lista_NFsPag"
Const VALOR_CONTRATOS = "Valor_Contratos"


'--- inicio dos properties get dos controles da tela

Public Property Get FrameTab() As Object
     Set FrameTab = objUserControl.Controls("FrameTab")
End Property

Public Property Get Modelo() As Object
     Set Modelo = objUserControl.Controls("Modelo")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get SoAtivos() As Object
     Set SoAtivos = objUserControl.Controls("SoAtivos")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get EntreCont() As Object
     Set EntreCont = objUserControl.Controls("EntreCont")
End Property

Public Property Get ApenasCont() As Object
     Set ApenasCont = objUserControl.Controls("ApenasCont")
End Property

Public Property Get FrameD1() As Object
     Set FrameD1 = objUserControl.Controls("FrameD1")
End Property

Public Property Get DataContAte() As Object
     Set DataContAte = objUserControl.Controls("DataContAte")
End Property

Public Property Get UpDownDataContDe() As Object
     Set UpDownDataContDe = objUserControl.Controls("UpDownDataContDe")
End Property

Public Property Get DataContDe() As Object
     Set DataContDe = objUserControl.Controls("DataContDe")
End Property

Public Property Get UpDownDataContAte() As Object
     Set UpDownDataContAte = objUserControl.Controls("UpDownDataContAte")
End Property

Public Property Get ApenasQualifCont() As Object
     Set ApenasQualifCont = objUserControl.Controls("ApenasQualifCont")
End Property

Public Property Get ApenasDiasCont() As Object
     Set ApenasDiasCont = objUserControl.Controls("ApenasDiasCont")
End Property

Public Property Get EntreQualifContDe() As Object
     Set EntreQualifContDe = objUserControl.Controls("EntreQualifContDe")
End Property

Public Property Get EntreQualifContAte() As Object
     Set EntreQualifContAte = objUserControl.Controls("EntreQualifContAte")
End Property

Public Property Get EntreDiasContDe() As Object
     Set EntreDiasContDe = objUserControl.Controls("EntreDiasContDe")
End Property

Public Property Get EntreDiasContAte() As Object
     Set EntreDiasContAte = objUserControl.Controls("EntreDiasContAte")
End Property

Public Property Get FaixaDataCont() As Object
     Set FaixaDataCont = objUserControl.Controls("FaixaDataCont")
End Property

Public Property Get FrameD4() As Object
     Set FrameD4 = objUserControl.Controls("FrameD4")
End Property

Public Property Get EntreQualifPContAte() As Object
     Set EntreQualifPContAte = objUserControl.Controls("EntreQualifPContAte")
End Property

Public Property Get EntreQualifPContDe() As Object
     Set EntreQualifPContDe = objUserControl.Controls("EntreQualifPContDe")
End Property

Public Property Get EntreDiasPContDe() As Object
     Set EntreDiasPContDe = objUserControl.Controls("EntreDiasPContDe")
End Property

Public Property Get EntreDiasPContAte() As Object
     Set EntreDiasPContAte = objUserControl.Controls("EntreDiasPContAte")
End Property

Public Property Get ApenasQualifPCont() As Object
     Set ApenasQualifPCont = objUserControl.Controls("ApenasQualifPCont")
End Property

Public Property Get ApenasDiasPCont() As Object
     Set ApenasDiasPCont = objUserControl.Controls("ApenasDiasPCont")
End Property

Public Property Get DataPContAte() As Object
     Set DataPContAte = objUserControl.Controls("DataPContAte")
End Property

Public Property Get UpDownDataPContDe() As Object
     Set UpDownDataPContDe = objUserControl.Controls("UpDownDataPContDe")
End Property

Public Property Get DataPContDe() As Object
     Set DataPContDe = objUserControl.Controls("DataPContDe")
End Property

Public Property Get UpDownDataPContAte() As Object
     Set UpDownDataPContAte = objUserControl.Controls("UpDownDataPContAte")
End Property

Public Property Get ApenasPCont() As Object
     Set ApenasPCont = objUserControl.Controls("ApenasPCont")
End Property

Public Property Get EntrePCont() As Object
     Set EntrePCont = objUserControl.Controls("EntrePCont")
End Property

Public Property Get FaixaDataPCont() As Object
     Set FaixaDataPCont = objUserControl.Controls("FaixaDataPCont")
End Property

Public Property Get FrameD2() As Object
     Set FrameD2 = objUserControl.Controls("FrameD2")
End Property

Public Property Get EntreQualifPCompAte() As Object
     Set EntreQualifPCompAte = objUserControl.Controls("EntreQualifPCompAte")
End Property

Public Property Get EntreQualifPCompDe() As Object
     Set EntreQualifPCompDe = objUserControl.Controls("EntreQualifPCompDe")
End Property

Public Property Get EntreDiasPCompDe() As Object
     Set EntreDiasPCompDe = objUserControl.Controls("EntreDiasPCompDe")
End Property

Public Property Get EntreDiasPCompAte() As Object
     Set EntreDiasPCompAte = objUserControl.Controls("EntreDiasPCompAte")
End Property

Public Property Get ApenasQualifPComp() As Object
     Set ApenasQualifPComp = objUserControl.Controls("ApenasQualifPComp")
End Property

Public Property Get ApenasDiasPComp() As Object
     Set ApenasDiasPComp = objUserControl.Controls("ApenasDiasPComp")
End Property

Public Property Get DataPCompAte() As Object
     Set DataPCompAte = objUserControl.Controls("DataPCompAte")
End Property

Public Property Get UpDownDataPCompDe() As Object
     Set UpDownDataPCompDe = objUserControl.Controls("UpDownDataPCompDe")
End Property

Public Property Get DataPCompDe() As Object
     Set DataPCompDe = objUserControl.Controls("DataPCompDe")
End Property

Public Property Get UpDownDataPCompAte() As Object
     Set UpDownDataPCompAte = objUserControl.Controls("UpDownDataPCompAte")
End Property

Public Property Get ApenasPComp() As Object
     Set ApenasPComp = objUserControl.Controls("ApenasPComp")
End Property

Public Property Get EntrePComp() As Object
     Set EntrePComp = objUserControl.Controls("EntrePComp")
End Property

Public Property Get FaixaDataPComp() As Object
     Set FaixaDataPComp = objUserControl.Controls("FaixaDataPComp")
End Property

Public Property Get EntreUComp() As Object
     Set EntreUComp = objUserControl.Controls("EntreUComp")
End Property

Public Property Get ApenasUComp() As Object
     Set ApenasUComp = objUserControl.Controls("ApenasUComp")
End Property

Public Property Get FrameD3() As Object
     Set FrameD3 = objUserControl.Controls("FrameD3")
End Property

Public Property Get DataUCompAte() As Object
     Set DataUCompAte = objUserControl.Controls("DataUCompAte")
End Property

Public Property Get UpDownDataUCompDe() As Object
     Set UpDownDataUCompDe = objUserControl.Controls("UpDownDataUCompDe")
End Property

Public Property Get DataUCompDe() As Object
     Set DataUCompDe = objUserControl.Controls("DataUCompDe")
End Property

Public Property Get UpDownDataUCompAte() As Object
     Set UpDownDataUCompAte = objUserControl.Controls("UpDownDataUCompAte")
End Property

Public Property Get EntreQualifUCompDe() As Object
     Set EntreQualifUCompDe = objUserControl.Controls("EntreQualifUCompDe")
End Property

Public Property Get EntreQualifUCompAte() As Object
     Set EntreQualifUCompAte = objUserControl.Controls("EntreQualifUCompAte")
End Property

Public Property Get EntreDiasUCompDe() As Object
     Set EntreDiasUCompDe = objUserControl.Controls("EntreDiasUCompDe")
End Property

Public Property Get EntreDiasUCompAte() As Object
     Set EntreDiasUCompAte = objUserControl.Controls("EntreDiasUCompAte")
End Property

Public Property Get ApenasQualifUComp() As Object
     Set ApenasQualifUComp = objUserControl.Controls("ApenasQualifUComp")
End Property

Public Property Get ApenasDiasuComp() As Object
     Set ApenasDiasuComp = objUserControl.Controls("ApenasDiasuComp")
End Property

Public Property Get FaixaDataUComp() As Object
     Set FaixaDataUComp = objUserControl.Controls("FaixaDataUComp")
End Property

Public Property Get ApenasNComp() As Object
     Set ApenasNComp = objUserControl.Controls("ApenasNComp")
End Property

Public Property Get EntreNComp() As Object
     Set EntreNComp = objUserControl.Controls("EntreNComp")
End Property

Public Property Get FrameD5() As Object
     Set FrameD5 = objUserControl.Controls("FrameD5")
End Property

Public Property Get EntreDiasNCompDe() As Object
     Set EntreDiasNCompDe = objUserControl.Controls("EntreDiasNCompDe")
End Property

Public Property Get EntreDiasNCompAte() As Object
     Set EntreDiasNCompAte = objUserControl.Controls("EntreDiasNCompAte")
End Property

Public Property Get ApenasDiasNComp() As Object
     Set ApenasDiasNComp = objUserControl.Controls("ApenasDiasNComp")
End Property

Public Property Get DataNCompAte() As Object
     Set DataNCompAte = objUserControl.Controls("DataNCompAte")
End Property

Public Property Get UpDownDataNCompDe() As Object
     Set UpDownDataNCompDe = objUserControl.Controls("UpDownDataNCompDe")
End Property

Public Property Get DataNCompDe() As Object
     Set DataNCompDe = objUserControl.Controls("DataNCompDe")
End Property

Public Property Get UpDownDataNCompAte() As Object
     Set UpDownDataNCompAte = objUserControl.Controls("UpDownDataNCompAte")
End Property

Public Property Get FaixaDataNComp() As Object
     Set FaixaDataNComp = objUserControl.Controls("FaixaDataNComp")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get TiposCliente() As Object
     Set TiposCliente = objUserControl.Controls("TiposCliente")
End Property

Public Property Get EmailValido() As Object
     Set EmailValido = objUserControl.Controls("EmailValido")
End Property

Public Property Get RespCallCenter() As Object
     Set RespCallCenter = objUserControl.Controls("RespCallCenter")
End Property

Public Property Get FrameCategoriaCliente() As Object
     Set FrameCategoriaCliente = objUserControl.Controls("FrameCategoriaCliente")
End Property

Public Property Get CategoriaCliente() As Object
     Set CategoriaCliente = objUserControl.Controls("CategoriaCliente")
End Property

Public Property Get CategoriaClienteDe() As Object
     Set CategoriaClienteDe = objUserControl.Controls("CategoriaClienteDe")
End Property

Public Property Get CategoriaClienteTodas() As Object
     Set CategoriaClienteTodas = objUserControl.Controls("CategoriaClienteTodas")
End Property

Public Property Get CategoriaClienteAte() As Object
     Set CategoriaClienteAte = objUserControl.Controls("CategoriaClienteAte")
End Property

Public Property Get LabelCategoriaClienteAte() As Object
     Set LabelCategoriaClienteAte = objUserControl.Controls("LabelCategoriaClienteAte")
End Property

Public Property Get LabelCategoriaClienteDe() As Object
     Set LabelCategoriaClienteDe = objUserControl.Controls("LabelCategoriaClienteDe")
End Property

Public Property Get LabelCategoriaCliente() As Object
     Set LabelCategoriaCliente = objUserControl.Controls("LabelCategoriaCliente")
End Property

Public Property Get LabelModelo() As Object
     Set LabelModelo = objUserControl.Controls("LabelModelo")
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

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Selecionado() As Object
     Set Selecionado = objUserControl.Controls("Selecionado")
End Property

Public Property Get BotaoMarcar() As Object
     Set BotaoMarcar = objUserControl.Controls("BotaoMarcar")
End Property

Public Property Get BotaoDesmarcar() As Object
     Set BotaoDesmarcar = objUserControl.Controls("BotaoDesmarcar")
End Property

Public Property Get GridClientes() As Object
     Set GridClientes = objUserControl.Controls("GridClientes")
End Property

Public Property Get ValorTotalCompras() As Object
     Set ValorTotalCompras = objUserControl.Controls("ValorTotalCompras")
End Property

Public Property Get DataUContato() As Object
     Set DataUContato = objUserControl.Controls("DataUContato")
End Property

Public Property Get DataPCompra() As Object
     Set DataPCompra = objUserControl.Controls("DataPCompra")
End Property

Public Property Get DataUCompra() As Object
     Set DataUCompra = objUserControl.Controls("DataUCompra")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get OpcoesTela() As Object
     Set OpcoesTela = objUserControl.Controls("OpcoesTela")
End Property

Public Property Get OpcaoPadrao() As Object
     Set OpcaoPadrao = objUserControl.Controls("OpcaoPadrao")
End Property

Public Property Get LabelOpcao() As Object
     Set LabelOpcao = objUserControl.Controls("LabelOpcao")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
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

Public Property Get IgnoraJaEnviados() As Object
     Set IgnoraJaEnviados = objUserControl.Controls("IgnoraJaEnviados")
End Property

'--- fim dos properties get dos controles da tela

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Contato com Clientes - Por Email"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContatoClientePorEmail"

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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Set gobjContatoCliAnt = Nothing
    
    Set gcolClientesEnv = Nothing
    Set gcolFiliaisEnv = Nothing
    Set gcolEnderecosEnv = Nothing
    Set gcolEstatisticasEnv = Nothing

    Set gcolClientes = Nothing
    Set gcolFiliais = Nothing
    Set gcolEnderecos = Nothing
    Set gcolEstatisticas = Nothing
    Set gcolModelos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objGridClientes = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200035)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objTela As Object
Dim sDigDiscExt As String
Dim colCobranca As New Collection
Dim objModelo As ClassCobrancaEmailPadrao
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    gbTrazendoDados = False

    Set objGridClientes = New AdmGrid
    
    Set gcolClientes = New Collection
    Set gcolFiliais = New Collection
    Set gcolEnderecos = New Collection
    Set gcolEstatisticas = New Collection
    Set gcolModelos = New Collection

    Set gobjContatoCliAnt = New ClassContatoCliSel
    
    lErro = Inicializa_GridClientes(objGridClientes)
    If lErro <> SUCESSO Then gError 200036

    lErro = Carrega_ComboCategoriaCliente(CategoriaCliente)
    If lErro <> SUCESSO Then gError 200264

    lErro = Carrega_Usuarios
    If lErro <> SUCESSO Then gError 200265
    
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    
    'Le os modelos válidos para o atraso em questão
    lErro = CF("CobrancaEmailPadrao_Le_Tipo", colCobranca, TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE)
    If lErro <> SUCESSO Then gError 189429
    
    Set gcolModelos = colCobranca
    
    If gcolModelos.Count = 0 Then gError 200266
        
    'Carrega a Combo com os Dados da Colecao
    Carta.Clear
    For Each objModelo In colCobranca
        Carta.AddItem (objModelo.lCodigo & SEPARADOR & objModelo.sDescricao)
        Carta.ItemData(Carta.NewIndex) = objModelo.lCodigo
    Next
    
    'Carrega a Combo com os Dados da Colecao
    Modelo.Clear
    For Each objModelo In colCobranca
        Modelo.AddItem (objModelo.lCodigo & SEPARADOR & objModelo.sDescricao)
        Modelo.ItemData(Modelo.NewIndex) = objModelo.lCodigo
    Next

    'Guarda em objTela os dados dessa tela
    Set objTela = Me
    
    lErro = CF("Carrega_OpcoesTela", objTela, True)
    If lErro <> SUCESSO Then gError 200037
    
'    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
'    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
'    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)

    If FaixaDataCont.Value Then
        Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
        Call Limpa_FrameD1(FRAMED_ENTRE)
    End If
    If FaixaDataPComp.Value Then
        Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
        Call Limpa_FrameD2(FRAMED_ENTRE)
    End If
    If FaixaDataUComp.Value Then
        Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
        Call Limpa_FrameD3(FRAMED_ENTRE)
    End If
    If FaixaDataPCont.Value Then
        Call FrameD_Enabled(FrameD4, FRAMED_FAIXA)
        Call Limpa_FrameD4(FRAMED_ENTRE)
    End If
    If FaixaDataNComp.Value Then
        Call FrameD_Enabled(FrameD5, FRAMED_FAIXA)
        Call Limpa_FrameD5(FRAMED_ENTRE)
    End If
    
    If EntreCont.Value Then
        Call FrameD_Enabled(FrameD1, FRAMED_ENTRE)
        Call Limpa_FrameD1(FRAMED_ENTRE)
    End If
    If EntrePComp.Value Then
        Call FrameD_Enabled(FrameD2, FRAMED_ENTRE)
        Call Limpa_FrameD2(FRAMED_ENTRE)
    End If
    If EntreUComp.Value Then
        Call FrameD_Enabled(FrameD3, FRAMED_ENTRE)
        Call Limpa_FrameD3(FRAMED_ENTRE)
    End If
    If EntrePCont.Value Then
        Call FrameD_Enabled(FrameD4, FRAMED_ENTRE)
        Call Limpa_FrameD4(FRAMED_ENTRE)
    End If
    If EntreNComp.Value Then
        Call FrameD_Enabled(FrameD5, FRAMED_ENTRE)
        Call Limpa_FrameD5(FRAMED_ENTRE)
    End If
    
    If ApenasCont.Value Then
        Call FrameD_Enabled(FrameD1, FRAMED_APENAS)
        Call Limpa_FrameD1(FRAMED_ENTRE)
    End If
    If ApenasPComp.Value Then
        Call FrameD_Enabled(FrameD2, FRAMED_APENAS)
        Call Limpa_FrameD2(FRAMED_ENTRE)
    End If
    If ApenasUComp.Value Then
        Call FrameD_Enabled(FrameD3, FRAMED_APENAS)
        Call Limpa_FrameD3(FRAMED_ENTRE)
    End If
    If ApenasPCont.Value Then
        Call FrameD_Enabled(FrameD4, FRAMED_APENAS)
        Call Limpa_FrameD4(FRAMED_ENTRE)
    End If
    If ApenasNComp.Value Then
        Call FrameD_Enabled(FrameD5, FRAMED_APENAS)
        Call Limpa_FrameD5(FRAMED_ENTRE)
    End If
    
    RespCallCenter.Text = gsUsuario
    
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", 255, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 200224
    
    For Each objCodigoNome In colCodigoDescricao
        TiposCliente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        TiposCliente.ItemData(TiposCliente.NewIndex) = objCodigoNome.iCodigo
        TiposCliente.Selected(TiposCliente.NewIndex) = True
    Next
       
    iFrameAtual = TAB_SELECAO
    iAlterado = 0
    iAlteradoFiltro = 0

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 200036, 200037, 200224, 200264, 200265
            'erros tratados nas rotinas chamadas
            
        Case 200266
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_MODELOEMAIL_VALIDO_TELA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200038)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200039)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Selecao_Memoria(ByVal objContatoCliSel As ClassContatoCliSel) As Long

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim iLinha As Integer

On Error GoTo Erro_Move_Selecao_Memoria

    'If Len(Trim(RespCallCenter.Text)) = 0 Then gError 200040
    
    objContatoCliSel.lModeloForcado = LCodigo_Extrai(Modelo.Text)
    
    If SoAtivos.Value = vbChecked Then
        objContatoCliSel.iSoAtivos = MARCADO
    Else
        objContatoCliSel.iSoAtivos = DESMARCADO
    End If
    
    If EmailValido.Value = vbChecked Then
        objContatoCliSel.iSoComEmailValido = MARCADO
    End If
    
    objContatoCliSel.sRespCallCenter = RespCallCenter.Text

    If FaixaDataCont.Value Then
        objContatoCliSel.dtDataContDe = StrParaDate(DataContDe.Text)
        objContatoCliSel.dtDataContAte = StrParaDate(DataContAte.Text)
    End If
    
    If IgnoraJaEnviados.Value = vbChecked Then
        objContatoCliSel.iIgnoraJaEnviados = MARCADO
    Else
        objContatoCliSel.iIgnoraJaEnviados = DESMARCADO
    End If

    If ApenasCont.Value Then
    
        If ApenasQualifCont.ListIndex = -1 Then gError 200041
    
        lErro = Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasCont.Text), ApenasQualifCont.ItemData(ApenasQualifCont.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataContDe = dtDataDe
        objContatoCliSel.dtDataContAte = dtDataAte
    End If
    
    If EntreCont.Value Then
        
        If EntreQualifContDe.ListIndex = -1 Then gError 200042
        If EntreQualifContAte.ListIndex = -1 Then gError 200043
        
        lErro = Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasContDe), EntreQualifContDe.ItemData(EntreQualifContDe.ListIndex), StrParaInt(EntreDiasContAte), EntreQualifContAte.ItemData(EntreQualifContAte.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataContDe = dtDataDe
        objContatoCliSel.dtDataContAte = dtDataAte
    End If
    
   If FaixaDataPCont.Value Then
        objContatoCliSel.dtDataPContDe = StrParaDate(DataPContDe.Text)
        objContatoCliSel.dtDataPContAte = StrParaDate(DataPContAte.Text)
    End If

    If ApenasPCont.Value Then
    
        If ApenasQualifPCont.ListIndex = -1 Then gError 200041
    
        lErro = Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasPCont.Text), ApenasQualifPCont.ItemData(ApenasQualifPCont.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataPContDe = dtDataDe
        objContatoCliSel.dtDataPContAte = dtDataAte
    End If
    
    If EntrePCont.Value Then
        
        If EntreQualifPContDe.ListIndex = -1 Then gError 200042
        If EntreQualifPContAte.ListIndex = -1 Then gError 200043
        
        lErro = Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasPContDe), EntreQualifPContDe.ItemData(EntreQualifPContDe.ListIndex), StrParaInt(EntreDiasPContAte), EntreQualifPContAte.ItemData(EntreQualifPContAte.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataPContDe = dtDataDe
        objContatoCliSel.dtDataPContAte = dtDataAte
    End If
    
    If FaixaDataNComp.Value Then
        objContatoCliSel.dtDataNCompDe = StrParaDate(DataNCompDe.Text)
        objContatoCliSel.dtDataNCompAte = StrParaDate(DataNCompAte.Text)
    End If

    If ApenasNComp.Value Then
    
        lErro = Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasNComp.Text), DATA_ATRAS)
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataNCompDe = dtDataDe
        objContatoCliSel.dtDataNCompAte = dtDataAte
    End If
    
    If EntreNComp.Value Then
        
        lErro = Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasNCompDe), DATA_ATRAS, StrParaInt(EntreDiasNCompAte), DATA_ATRAS)
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataNCompDe = dtDataDe
        objContatoCliSel.dtDataNCompAte = dtDataAte
    End If

    If FaixaDataPComp.Value Then
        objContatoCliSel.dtDataPCompDe = StrParaDate(DataPCompDe.Text)
        objContatoCliSel.dtDataPCompAte = StrParaDate(DataPCompAte.Text)
    End If

    If ApenasPComp.Value Then
    
        If ApenasQualifPComp.ListIndex = -1 Then gError 200044

        lErro = Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasPComp.Text), ApenasQualifPComp.ItemData(ApenasQualifPComp.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataPCompDe = dtDataDe
        objContatoCliSel.dtDataPCompAte = dtDataAte
    End If
    
    If EntrePComp.Value Then
    
        If EntreQualifPCompDe.ListIndex = -1 Then gError 200045
        If EntreQualifPCompAte.ListIndex = -1 Then gError 200046

        lErro = Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasPCompDe), EntreQualifPCompDe.ItemData(EntreQualifPCompDe.ListIndex), StrParaInt(EntreDiasPCompAte), EntreQualifPCompAte.ItemData(EntreQualifPCompAte.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataPCompDe = dtDataDe
        objContatoCliSel.dtDataPCompAte = dtDataAte
    End If
    
    If FaixaDataUComp.Value Then
        objContatoCliSel.dtDataUCompDe = StrParaDate(DataUCompDe.Text)
        objContatoCliSel.dtDataUCompAte = StrParaDate(DataUCompAte.Text)
    End If

    If ApenasUComp.Value Then
    
        If ApenasQualifUComp.ListIndex = -1 Then gError 200047

        lErro = Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasuComp.Text), ApenasQualifUComp.ItemData(ApenasQualifUComp.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataUCompDe = dtDataDe
        objContatoCliSel.dtDataUCompAte = dtDataAte
    End If
    
    If EntreUComp.Value Then
        
        If EntreQualifUCompDe.ListIndex = -1 Then gError 200048
        If EntreQualifUCompAte.ListIndex = -1 Then gError 200049
        
        lErro = Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasUCompDe), EntreQualifUCompDe.ItemData(EntreQualifUCompDe.ListIndex), StrParaInt(EntreDiasUCompAte.Text), EntreQualifUCompAte.ItemData(EntreQualifUCompAte.ListIndex))
        If lErro <> SUCESSO Then gError 200033
        
        objContatoCliSel.dtDataUCompDe = dtDataDe
        objContatoCliSel.dtDataUCompAte = dtDataAte
    End If
    
    If objContatoCliSel.dtDataUCompAte <> DATA_NULA And objContatoCliSel.dtDataUCompDe <> DATA_NULA Then
        If objContatoCliSel.dtDataUCompDe > objContatoCliSel.dtDataUCompAte Then gError 200050
    End If
    
    If objContatoCliSel.dtDataContDe <> DATA_NULA And objContatoCliSel.dtDataContAte <> DATA_NULA Then
        If objContatoCliSel.dtDataContDe > objContatoCliSel.dtDataContAte Then gError 200051
    End If
    
    If objContatoCliSel.dtDataPCompDe <> DATA_NULA And objContatoCliSel.dtDataPCompAte <> DATA_NULA Then
        If objContatoCliSel.dtDataPCompDe > objContatoCliSel.dtDataPCompAte Then gError 200052
    End If
    
    If CategoriaClienteTodas.Value = vbChecked Then
        objContatoCliSel.sCategoria = ""
        objContatoCliSel.sCategoriaDe = ""
        objContatoCliSel.sCategoriaAte = ""
    Else
        If CategoriaCliente.Text = "" Then gError 200053
        objContatoCliSel.sCategoria = CategoriaCliente.Text
        objContatoCliSel.sCategoriaDe = CategoriaClienteDe.Text
        objContatoCliSel.sCategoriaAte = CategoriaClienteAte.Text
    End If
    
    For iLinha = 0 To TiposCliente.ListCount - 1
        
        If Not TiposCliente.Selected(iLinha) Then
            objContatoCliSel.colTiposNaoConsiderar.Add TiposCliente.ItemData(iLinha)
        End If
    
    Next
    
    If objContatoCliSel.sCategoriaDe > objContatoCliSel.sCategoriaAte Then gError 200054

    lErro = CF2(Me, "CONTCLIE_Move_Selecao_Memoria", objContatoCliSel)
    If lErro <> SUCESSO Then gError 200988

    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
    
        Case 200040
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIOCOBRADOR_NAO_PREENCHIDO", gErr)
                   
        Case 200041 To 200049
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)
    
        Case 200050 To 200052
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            
        Case 200053
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
            
        Case 200054
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ITEM_INICIAL_MAIOR", gErr)
            
        Case 200988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200055)

    End Select

    Exit Function

End Function

Function Datas_Trata_Entre(dtDataDe As Date, dtDataAte As Date, ByVal iNumDias1 As Integer, ByVal iData1 As Integer, ByVal iNumDias2 As Integer, ByVal iData2 As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Datas_Trata_Entre

    Select Case iData1
    
        Case DATA_AFRENTE
            dtDataDe = DateAdd("d", iNumDias1, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataDe = DateAdd("d", -iNumDias1, gdtDataAtual)
            
        Case DATA_IGNORAR
            If iNumDias1 <> 0 Then gError 200031
            dtDataDe = DATA_NULA
            
        Case Else
            gError 200056

    End Select
    
    Select Case iData2
    
        Case DATA_AFRENTE
            dtDataAte = DateAdd("d", iNumDias2, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataAte = DateAdd("d", -iNumDias2, gdtDataAtual)
        
        Case DATA_IGNORAR
            If iNumDias2 <> 0 Then gError 200031
            dtDataAte = DATA_NULA
        
        Case Else
            gError 200057

    End Select
   
    Datas_Trata_Entre = SUCESSO

    Exit Function

Erro_Datas_Trata_Entre:

    Datas_Trata_Entre = gErr

    Select Case gErr
    
        Case 200031
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_IGNORA_COM_VALOR", gErr)
    
        Case 200056, 200057
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200058)

    End Select

    Exit Function

End Function

Function Datas_Trata_Apenas(dtDataDe As Date, dtDataAte As Date, ByVal iNumDias As Integer, ByVal iData As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Datas_Trata_Apenas

    Select Case iData
    
        Case DATA_AFRENTE
            dtDataDe = gdtDataAtual
            dtDataAte = DateAdd("d", iNumDias, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataDe = DateAdd("d", -iNumDias, gdtDataAtual)
            dtDataAte = gdtDataAtual
            
        Case Else
            gError 200059

    End Select
   
    Datas_Trata_Apenas = SUCESSO

    Exit Function

Erro_Datas_Trata_Apenas:

    Datas_Trata_Apenas = gErr

    Select Case gErr
    
        Case 200059
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200060)

    End Select

    Exit Function

End Function

Function Trata_Selecao(ByVal objContatoCliSel As ClassContatoCliSel) As Long

Dim lErro As Long
Dim colFiliais As New Collection
Dim colEnderecos As New Collection
Dim colClienteEst As New Collection
Dim objEmailConfig As New ClassEmailConfig

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    objEmailConfig.sUsuario = gsUsuario
    
    lErro = CF("EmailConfig_Le", objEmailConfig)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202829
    
    If lErro <> SUCESSO Then gError 202830
    
    lErro = CF("ContatoCli_Selecao_Le", objContatoCliSel, colFiliais, colEnderecos, colClienteEst)
    If lErro <> SUCESSO Then gError 200061
    
    If colFiliais.Count = 0 Then gError 200062
    
    lErro = Preenche_GridCliente(objContatoCliSel, colFiliais, colEnderecos, colClienteEst)
    If lErro <> SUCESSO Then gError 200063
    
    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 200061, 200063, 202829
        
        Case 200062
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_COBRANCA_SEM_CLIENTES", gErr)
            
        Case 202830
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCONFIG_NAO_CADASTRADO", gErr, objEmailConfig.sUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200064)

    End Select

    Exit Function

End Function

Function Preenche_GridCliente(ByVal objContatoCliSel As ClassContatoCliSel, ByVal colFiliais As Collection, ByVal colEnderecos As Collection, ByVal colEstatisticas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilial As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim colCliente As New Collection
Dim objCliente As ClassCliente
Dim bAchou As Boolean
Dim objClieEst As ClassFilialClienteEst
Dim iPos As Integer
Dim objModelo As ClassCobrancaEmailPadrao
Dim objTela As Object
Dim sEmail As String

On Error GoTo Erro_Preenche_GridCliente

    Set gcolClientes = New Collection
    Set gcolEnderecos = New Collection
    Set gcolEstatisticas = New Collection
    Set gcolFiliais = New Collection
    
    gbTrazendoDados = True

    Call Grid_Limpa(objGridClientes)
    
    'Aumenta o número de linhas do grid se necessário
    If colFiliais.Count >= objGridClientes.objGrid.Rows Then
        Call Refaz_Grid(objGridClientes, colFiliais.Count)
    End If

    iIndice = 0
    For Each objFilial In colFiliais

        iIndice = iIndice + 1

        Set objClieEst = colEstatisticas(iIndice)
        Set objEndereco = colEnderecos(iIndice)
        
        bAchou = False
        iPos = 0
        For Each objCliente In colCliente
            iPos = iPos + 1
            If objCliente.lCodigo = objFilial.lCodCliente Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objCliente = New ClassCliente
            
            objCliente.lCodigo = objFilial.lCodCliente
        
            'le o cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 200065
        
        End If
        
        'Força o modelo de email que estará como padrão
        If objContatoCliSel.lModeloForcado <> 0 Then
            bAchou = False
            For Each objModelo In gcolModelos
                If objModelo.lCodigo = objContatoCliSel.lModeloForcado Then
                    bAchou = True
                    Exit For
                End If
            Next
        Else
            Set objModelo = gcolModelos.Item(1)
        End If
        
        If objModelo.lCodigo <> 0 Then
            GridClientes.TextMatrix(iIndice, iGrid_Carta_Col) = objModelo.lCodigo & SEPARADOR & objModelo.sDescricao
        End If
        
        sEmail = objEndereco.sEmail
        If Len(Trim(objEndereco.sEmail2)) > 0 Then
            If sEmail <> "" Then sEmail = sEmail & ";"
            sEmail = sEmail & objEndereco.sEmail2
        End If
        GridClientes.TextMatrix(iIndice, iGrid_Email_Col) = sEmail

        GridClientes.TextMatrix(iIndice, iGrid_Selecionado_Col) = CStr(DESMARCADO)
        
        GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        GridClientes.TextMatrix(iIndice, iGrid_Filial_Col) = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
        GridClientes.TextMatrix(iIndice, iGrid_ValorCompra_Col) = Format(objClieEst.dValorAcumuladoCompras, "STANDARD")
        
        If objClieEst.dtDataUltimoContato <> DATA_NULA Then
            GridClientes.TextMatrix(iIndice, iGrid_DataUContato_Col) = Format(objClieEst.dtDataUltimoContato, "dd/mm/yyyy")
        Else
            GridClientes.TextMatrix(iIndice, iGrid_DataUContato_Col) = ""
        End If
        
        If objClieEst.dtDataUltimaCompra <> DATA_NULA Then
            GridClientes.TextMatrix(iIndice, iGrid_DataUCompra_Col) = Format(objClieEst.dtDataUltimaCompra, "dd/mm/yyyy")
        Else
            GridClientes.TextMatrix(iIndice, iGrid_DataUCompra_Col) = ""
        End If
        
        If objClieEst.dtDataPrimeiraCompra <> DATA_NULA Then
            GridClientes.TextMatrix(iIndice, iGrid_DataPCompra_Col) = Format(objClieEst.dtDataPrimeiraCompra, "dd/mm/yyyy")
        Else
            GridClientes.TextMatrix(iIndice, iGrid_DataPCompra_Col) = ""
        End If
        
        gcolClientes.Add objCliente
        gcolEnderecos.Add objEndereco
        gcolEstatisticas.Add objClieEst
        gcolFiliais.Add objFilial
        
    Next
           
    objGridClientes.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridClientes)
    
    Set objTela = Me
    
    For iIndice = 1 To objGridClientes.iLinhasExistentes
    
        giLinhaAtual = iIndice
        
        lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objModelo)
        If lErro <> SUCESSO Then gError 187034
    
        GridClientes.TextMatrix(iIndice, iGrid_CC_Col) = objModelo.sCCValor
        GridClientes.TextMatrix(iIndice, iGrid_Assunto_Col) = objModelo.sAssuntoValor
        GridClientes.TextMatrix(iIndice, iGrid_Mensagem_Col) = objModelo.sMensagemValor
        GridClientes.TextMatrix(iIndice, iGrid_Anexo_Col) = objModelo.sAnexoValor
        
        If Len(Trim(objModelo.sMensagemValor)) = 0 Then
            GridClientes.TextMatrix(iIndice, iGrid_Mensagem_Col) = "<Mensagem automático confome modelo '" & objModelo.sModelo & "'>"
        End If

    Next

    gbTrazendoDados = False

    Preenche_GridCliente = SUCESSO

    Exit Function

Erro_Preenche_GridCliente:

    gbTrazendoDados = False

    Preenche_GridCliente = gErr

    Select Case gErr
    
        Case 200065, 200066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200067)

    End Select

    Exit Function

End Function

Private Function Mostra_Dados_Email(ByVal iLinha As Integer) As Long

Dim objModelo As ClassCobrancaEmailPadrao
Dim lModelo As Long

    If iLinha <> 0 Then
        Cc.Text = GridClientes.TextMatrix(iLinha, iGrid_CC_Col)
        Assunto.Text = GridClientes.TextMatrix(iLinha, iGrid_Assunto_Col)
        Anexo.Text = GridClientes.TextMatrix(iLinha, iGrid_Anexo_Col)
        Mensagem.Text = GridClientes.TextMatrix(iLinha, iGrid_Mensagem_Col)
        Email.Text = GridClientes.TextMatrix(iLinha, iGrid_Email_Col)
        lModelo = LCodigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Carta_Col))
        Mensagem.Enabled = False
'        For Each objModelo In gcolModelos
'            If objModelo.lCodigo = lModelo Then
'                If Len(Trim(objModelo.sModelo)) > 0 Then
'                    Mensagem.Enabled = False
'                Else
'                    Mensagem.Enabled = True
'                End If
'                Exit For
'            End If
'        Next
        
    End If
    
    Mostra_Dados_Email = SUCESSO

End Function

Private Function Recolhe_Dados_Email(ByVal iLinha As Integer) As Long

    If iLinha <> 0 Then
        GridClientes.TextMatrix(iLinha, iGrid_CC_Col) = Cc.Text
        GridClientes.TextMatrix(iLinha, iGrid_Assunto_Col) = Assunto.Text
        GridClientes.TextMatrix(iLinha, iGrid_Mensagem_Col) = Mensagem.Text
        GridClientes.TextMatrix(iLinha, iGrid_Email_Col) = Email.Text
        GridClientes.TextMatrix(iLinha, iGrid_Anexo_Col) = Anexo.Text
    End If
    
    Recolhe_Dados_Email = SUCESSO

End Function

Private Function Preenche_Dados_Carta(ByVal lCarta As Long, ByVal lCartaAnt As Long, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objModelo As New ClassCobrancaEmailPadrao
Dim objTela As Object

On Error GoTo Erro_Preenche_Dados_Carta

    If iLinha <> 0 And lCarta <> 0 Then
    
        'Se trocou o tipo de carta
        If lCarta <> lCartaAnt Then

            For Each objModelo In gcolModelos
                If objModelo.lCodigo = lCarta Then
                    Exit For
                End If
            Next

            Set objTela = Me
            
            giLinhaAtual = iLinha
                
            lErro = CF("CobrancaEmailPadrao_Calcula_Regras", objTela, objModelo)
            If lErro <> SUCESSO Then gError 187036
        
            GridClientes.TextMatrix(iLinha, iGrid_CC_Col) = objModelo.sCCValor
            GridClientes.TextMatrix(iLinha, iGrid_Assunto_Col) = objModelo.sAssuntoValor
            GridClientes.TextMatrix(iLinha, iGrid_Mensagem_Col) = objModelo.sMensagemValor
            GridClientes.TextMatrix(iLinha, iGrid_Anexo_Col) = objModelo.sAnexoValor
            
            If Len(Trim(objModelo.sModelo)) > 0 Then
                GridClientes.TextMatrix(iLinha, iGrid_Mensagem_Col) = "<Mensagem confome modelo '" & objModelo.sModelo & "'>"
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
Dim objModelo As ClassCobrancaEmailPadrao
Dim objEnvioEmail As ClassEnvioDeEmail
Dim colEnvioEmail As New Collection
Dim sNomeArqParam As String
Dim vValor As Variant
Dim lModelo As Long

On Error GoTo Erro_BotaoEmail_Click
    
    Call Recolhe_Dados_Email(iLinhaAnt)
    
    For iLinha = 1 To objGridClientes.iLinhasExistentes
                   
        If StrParaInt(GridClientes.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
       
            If Len(Trim(GridClientes.TextMatrix(iLinha, iGrid_Email_Col))) = 0 Then gError 187038
    
            If InStr(1, GridClientes.TextMatrix(iLinha, iGrid_Email_Col), "@") = 0 Or InStr(1, GridClientes.TextMatrix(iLinha, iGrid_Email_Col), ".") = 0 Then gError 187038
    
            lErro = Verifica_Existencia_Arquivo(GridClientes.TextMatrix(iLinha, iGrid_Anexo_Col))
            If lErro <> SUCESSO Then gError 189413
    
        End If
    
    Next
      
    Set gcolClientesEnv = New Collection
    Set gcolEnderecosEnv = New Collection
    Set gcolFiliaisEnv = New Collection
    Set gcolEstatisticasEnv = New Collection
    
    For Each vValor In gcolClientes
        gcolClientesEnv.Add vValor
    Next
    
    For Each vValor In gcolEnderecos
        gcolEnderecosEnv.Add vValor
    Next
    
    For Each vValor In gcolFiliais
        gcolFiliaisEnv.Add vValor
    Next
    
    For Each vValor In gcolEstatisticas
        gcolEstatisticasEnv.Add vValor
    Next
    
    For iLinha = 1 To objGridClientes.iLinhasExistentes
                      
        If StrParaInt(GridClientes.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
        
            Set objEnvioEmail = New ClassEnvioDeEmail
            colEnvioEmail.Add objEnvioEmail
        
            objEnvioEmail.sCC = GridClientes.TextMatrix(iLinha, iGrid_CC_Col)
            objEnvioEmail.sEmail = GridClientes.TextMatrix(iLinha, iGrid_Email_Col)
            objEnvioEmail.sAssunto = GridClientes.TextMatrix(iLinha, iGrid_Assunto_Col)
            objEnvioEmail.sMensagem = GridClientes.TextMatrix(iLinha, iGrid_Mensagem_Col)
            objEnvioEmail.sAnexo = GridClientes.TextMatrix(iLinha, iGrid_Anexo_Col)
            objEnvioEmail.iLinha = iLinha
            
            Set objEnvioEmail.objTela = Me
            
            lModelo = LCodigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Carta_Col))
            
            For Each objModelo In gcolModelos
                If objModelo.lCodigo = lModelo Then
                    Exit For
                End If
            Next
                        
            objEnvioEmail.lClienteRelac = gcolClientes.Item(iLinha).lCodigo
            objEnvioEmail.lNumIntDocParc = 0
            objEnvioEmail.iGeraRelac = MARCADO
            objEnvioEmail.iFilialCliRelac = gcolFiliais.Item(iLinha).iCodFilial
            objEnvioEmail.sTextoRelac = objEnvioEmail.sMensagem
            
            objEnvioEmail.iTipoDoc = EMAIL_TIPODOC_CLIENTE
            objEnvioEmail.lNumIntDoc = gcolClientes.Item(iLinha).lCodigo
            objEnvioEmail.lNumIntDocModelo = objModelo.lNumIntDoc
            
            If InStr(1, objEnvioEmail.sTextoRelac, "<Mensagem automático confome modelo") <> 0 Then
                sMsgAux = ""
                If Len(Trim(objModelo.sModelo)) = 0 Then
                    sMsgAux = objEnvioEmail.sMensagem
                Else
                    sMsgAux = "Email confome modelo '" & objModelo.sModelo & "', enviado para o email '" & objEnvioEmail.sEmail & "', com cópia para '" & objEnvioEmail.sCC & "', assunto '" & objEnvioEmail.sAssunto & "' e contendo como anexo o arquivo '" & objEnvioEmail.sAnexo & "'."
                End If
                objEnvioEmail.sTextoRelac = sMsgAux
            End If
            
            objEnvioEmail.sModelo = objModelo.sModelo
            objEnvioEmail.sDe = objModelo.sDe
            objEnvioEmail.sNomeExibicao = objModelo.sNomeExibicao
        
        End If
       
    Next
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 196974

    lErro = CF("Rotina_Envia_Emails_Batch", sNomeArqParam, colEnvioEmail)
    If lErro <> SUCESSO Then gError 196975

    iAlterado = 0

    Exit Sub
    
Erro_BotaoEmail_Click:

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

Public Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim objContatoCli As New ClassContatoCliSel

On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_SELECAO Then
    
        'Valida a seleção
        lErro = Move_Selecao_Memoria(objContatoCli)
        If lErro <> SUCESSO Then gError 200068
        
        If objContatoCli.dtDataUCompAte <> gobjContatoCliAnt.dtDataUCompAte Or _
            objContatoCli.dtDataUCompDe <> gobjContatoCliAnt.dtDataUCompDe Or _
            objContatoCli.dtDataContAte <> gobjContatoCliAnt.dtDataContAte Or _
            objContatoCli.dtDataContDe <> gobjContatoCliAnt.dtDataContDe Or _
            objContatoCli.dtDataPContAte <> gobjContatoCliAnt.dtDataPContAte Or _
            objContatoCli.dtDataPContDe <> gobjContatoCliAnt.dtDataPContDe Or _
            objContatoCli.dtDataPCompAte <> gobjContatoCliAnt.dtDataPCompAte Or _
            objContatoCli.dtDataPCompDe <> gobjContatoCliAnt.dtDataPCompDe Or _
            objContatoCli.dtDataNCompAte <> gobjContatoCliAnt.dtDataNCompAte Or _
            objContatoCli.dtDataNCompDe <> gobjContatoCliAnt.dtDataNCompDe Or _
            objContatoCli.sRespCallCenter <> gobjContatoCliAnt.sRespCallCenter Or _
            objContatoCli.sCategoria <> gobjContatoCliAnt.sCategoria Or _
            objContatoCli.sCategoriaDe <> gobjContatoCliAnt.sCategoriaDe Or _
            objContatoCli.sCategoriaAte <> gobjContatoCliAnt.sCategoriaAte Or _
            objContatoCli.iSoAtivos <> gobjContatoCliAnt.iSoAtivos Or _
            iAlteradoFiltro = REGISTRO_ALTERADO Then
            
            iAlteradoFiltro = 0
            
            lErro = Trata_Selecao(objContatoCli)
            If lErro <> SUCESSO Then gError 200069
            
            Set gobjContatoCliAnt = objContatoCli
            
        End If
    
    End If

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 200068, 200069

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200070)

    End Select

    Exit Sub

End Sub

Public Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameTab(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameTab(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Function Inicializa_GridClientes(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Vlr.Compra")
    objGrid.colColuna.Add ("Últ.Contato")
    objGrid.colColuna.Add ("Pri.Compra")
    objGrid.colColuna.Add ("Últ.Compra")
    objGrid.colColuna.Add ("Modelo")
    objGrid.colColuna.Add ("Email")
    objGrid.colColuna.Add ("CC")
    objGrid.colColuna.Add ("Assunto")
    objGrid.colColuna.Add ("Anexo")
    objGrid.colColuna.Add ("Mensagem")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (Selecionado.Name)
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (Filial.Name)
    objGrid.colCampo.Add (ValorTotalCompras.Name)
    objGrid.colCampo.Add (DataUContato.Name)
    objGrid.colCampo.Add (DataPCompra.Name)
    objGrid.colCampo.Add (DataUCompra.Name)
    objGrid.colCampo.Add (Carta.Name)
    objGrid.colCampo.Add (EmailGrid.Name)
    objGrid.colCampo.Add (CCGrid.Name)
    objGrid.colCampo.Add (AssuntoGrid.Name)
    objGrid.colCampo.Add (AnexoGrid.Name)
    objGrid.colCampo.Add (MensagemGrid.Name)

    'Colunas do Grid
    iGrid_Selecionado_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_Filial_Col = 3
    iGrid_ValorCompra_Col = 4
    iGrid_DataUContato_Col = 5
    iGrid_DataPCompra_Col = 6
    iGrid_DataUCompra_Col = 7
    iGrid_Carta_Col = 8
    iGrid_Email_Col = 9
    iGrid_CC_Col = 10
    iGrid_Assunto_Col = 11
    iGrid_Anexo_Col = 12
    iGrid_Mensagem_Col = 13

    objGrid.objGrid = GridClientes

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridClientes.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridClientes = SUCESSO

End Function

Public Sub GridClientes_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If
    
    colcolColecoes.Add gcolClientes
    colcolColecoes.Add gcolFiliais
    colcolColecoes.Add gcolEnderecos
    colcolColecoes.Add gcolEstatisticas
    
    Call Ordenacao_ClickGrid(objGridClientes, Nothing, colcolColecoes)
    
End Sub

Public Sub GridClientes_GotFocus()
    Call Grid_Recebe_Foco(objGridClientes)
End Sub

Public Sub GridClientes_EnterCell()
    Call Grid_Entrada_Celula(objGridClientes, iAlterado)
End Sub

Public Sub GridClientes_LeaveCell()
    Call Saida_Celula(objGridClientes)
End Sub

Public Sub GridClientes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If

End Sub

Public Sub GridClientes_RowColChange()

    Call Grid_RowColChange(objGridClientes)

    Call Recolhe_Dados_Email(iLinhaAnt)
    Call Mostra_Dados_Email(GridClientes.Row)
    
    iLinhaAnt = GridClientes.Row
    
End Sub

Public Sub GridClientes_Scroll()
    Call Grid_Scroll(objGridClientes)
End Sub

Public Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridClientes)
End Sub

Public Sub GridClientes_LostFocus()
    Call Grid_Libera_Foco(objGridClientes)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'OperacaoInsumos
        If objGridInt.objGrid.Name = GridClientes.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Carta_Col
                
                    lErro = Saida_Celula_Carta(objGridInt)
                    If lErro <> SUCESSO Then gError 187009

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 200074

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 200071 To 200073

        Case 200074
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200075)

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
    lCartaAnt = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Carta_Col))
    
    lErro = Preenche_Dados_Carta(lCarta, lCartaAnt, GridClientes.Row)
    If lErro <> SUCESSO Then gError 187012
    
    Call Mostra_Dados_Email(GridClientes.Row)
        
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

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 200090

    'Limpa a Tela
    lErro = Limpa_Tela_Cobranca
    If lErro <> SUCESSO Then gError 200091

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 200090, 200091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200092)

    End Select

End Sub

Function Limpa_Tela_Cobranca() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Cobranca
        
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridClientes)
    
    Set gobjContatoCliAnt = New ClassContatoCliSel
    
    sOpcaoAnt = ""
    
    SoAtivos.Value = vbUnchecked
    
    OpcoesTela.Text = ""
    RespCallCenter.Text = gsUsuario
    ApenasQualifCont.ListIndex = -1
    EntreQualifContDe.ListIndex = -1
    EntreQualifContAte.ListIndex = -1
    ApenasQualifPComp.ListIndex = -1
    EntreQualifPCompDe.ListIndex = -1
    EntreQualifPCompAte.ListIndex = -1
    ApenasQualifUComp.ListIndex = -1
    EntreQualifUCompDe.ListIndex = -1
    EntreQualifUCompAte.ListIndex = -1
    
    '#####################################
    'Inserido por Wagner
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    '#####################################
    
    FaixaDataUComp.Value = True
    FaixaDataCont.Value = True
    FaixaDataPComp.Value = True
    
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    
    Call Ordenacao_Limpa(objGridClientes)
    
    If iFrameAtual <> TAB_SELECAO Then
        'Torna Frame atual invisível
        FrameTab(TabStrip1.SelectedItem.Index).Visible = False
        iFrameAtual = TAB_SELECAO
        'Torna Frame atual visível
        FrameTab(iFrameAtual).Visible = True
        TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    End If
    
    Modelo.ListIndex = -1
    
    iAlterado = 0

    Limpa_Tela_Cobranca = SUCESSO

    Exit Function

Erro_Limpa_Tela_Cobranca:

    Limpa_Tela_Cobranca = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200093)

    End Select

    Exit Function

End Function

Private Sub UpDownData_DownClick(ByVal objDataMask As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    objDataMask.SetFocus

    If Len(objDataMask.ClipText) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 200094

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 200094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200095)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick(ByVal objDataMask As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    objDataMask.SetFocus

    If Len(Trim(objDataMask.ClipText)) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 200096

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 200096

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200097)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataUCompAte_DownClick()
    Call UpDownData_DownClick(DataUCompAte)
End Sub

Public Sub UpDownDataUCompAte_UpClick()
    Call UpDownData_UpClick(DataUCompAte)
End Sub

Public Sub UpDownDataUCompDe_DownClick()
    Call UpDownData_DownClick(DataUCompDe)
End Sub

Public Sub UpDownDataUCompDe_UpClick()
    Call UpDownData_UpClick(DataUCompDe)
End Sub

Public Sub UpDownDataContAte_DownClick()
    Call UpDownData_DownClick(DataContAte)
End Sub

Public Sub UpDownDataContAte_UpClick()
    Call UpDownData_UpClick(DataContAte)
End Sub

Public Sub UpDownDataContDe_DownClick()
    Call UpDownData_DownClick(DataContDe)
End Sub

Public Sub UpDownDataContDe_UpClick()
    Call UpDownData_UpClick(DataContDe)
End Sub

Public Sub UpDownDataPCompAte_DownClick()
    Call UpDownData_DownClick(DataPCompAte)
End Sub

Public Sub UpDownDataPCompAte_UpClick()
    Call UpDownData_UpClick(DataPCompAte)
End Sub

Public Sub UpDownDataPCompDe_DownClick()
    Call UpDownData_DownClick(DataPCompDe)
End Sub

Public Sub UpDownDataPCompDe_UpClick()
    Call UpDownData_UpClick(DataPCompDe)
End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 200098

    For Each objUsuarios In colUsuarios
        RespCallCenter.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 200098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200099)

    End Select

    Exit Function

End Function

Private Sub Marca_Desmarca(ByVal bFlag As Boolean)

Dim iIndice As Integer

    For iIndice = 1 To objGridClientes.iLinhasExistentes
    
        If bFlag Then
            GridClientes.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO
        Else
            GridClientes.TextMatrix(iIndice, iGrid_Selecionado_Col) = DESMARCADO
        End If
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridClientes)

End Sub

Public Sub BotaoDesmarcar_Click()
    Call Marca_Desmarca(False)
End Sub

Public Sub BotaoMarcar_Click()
    Call Marca_Desmarca(True)
End Sub

Public Sub OpcoesTela_Validate(Cancel As Boolean)
    'Se a opção não foi selecionada na combo => chama a função OpcoesTela_Click
    If OpcoesTela.ListIndex = -1 Then Call OpcoesTela_Click
End Sub

Public Function Gravar_Registro() As Long

Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Grava", objTela)
    If lErro <> SUCESSO Then gError 200100
    
    Call Limpa_Tela_Cobranca
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 200100
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200101)
        
    End Select
    
End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_BotaoExcluir_Click

    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Exclui", objTela)
    If lErro <> SUCESSO Then gError 200102
    
    'Call Limpa_Tela_Cobranca

    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 200102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200103)

    End Select

End Sub

Public Sub OpcoesTela_Click()
    
Dim lErro As Long
Dim objTela As Object
Dim iCancel As Integer

On Error GoTo Erro_OpcoesTela_Click

    Set objTela = Me
    
    If sOpcaoAnt <> OpcoesTela.Text Then
    
        iAtualizaTela = MARCADO
        
        'Trata o evento click da combo opções
        lErro = CF("OpcoesTela_Click", objTela)
        If lErro <> SUCESSO Then gError 200104
        
        sOpcaoAnt = OpcoesTela.Text
        
        'Se Frame selecionado foi o de seleção e é para atualizar o grid
        If TabStrip1.SelectedItem.Index = TAB_CLIENTE Then
        
            iCancel = bSGECancelDummy
        
            Call TabStrip1_BeforeClick(iCancel)
        
        End If
        
        Call CategoriaCliente_Validate(bSGECancelDummy)
        
        'Trata o evento click da combo opções
        lErro = CF("OpcoesTela_Click", objTela)
        If lErro <> SUCESSO Then gError 200104
        
        If FaixaDataCont.Value Then
            Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If FaixaDataPComp.Value Then
            Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If FaixaDataUComp.Value Then
            Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        
        If EntreCont.Value Then
            Call FrameD_Enabled(FrameD1, FRAMED_ENTRE)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If EntrePComp.Value Then
            Call FrameD_Enabled(FrameD2, FRAMED_ENTRE)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If EntreUComp.Value Then
            Call FrameD_Enabled(FrameD3, FRAMED_ENTRE)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        
        If ApenasCont.Value Then
            Call FrameD_Enabled(FrameD1, FRAMED_APENAS)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If ApenasPComp.Value Then
            Call FrameD_Enabled(FrameD2, FRAMED_APENAS)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        If ApenasUComp.Value Then
            Call FrameD_Enabled(FrameD3, FRAMED_APENAS)
            Call Limpa_FrameD1(FRAMED_ENTRE)
        End If
        
    End If
    
    Exit Sub

Erro_OpcoesTela_Click:

    Select Case gErr

        Case 200104
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200105)

    End Select

End Sub

Private Sub FrameD_Enabled(objFrame As Object, ByVal iIndice As Integer)

Dim iIndiceAux As Integer

    For iIndiceAux = 0 To 2
        If iIndiceAux = iIndice Then
            objFrame(iIndiceAux).Enabled = True
        Else
            objFrame(iIndiceAux).Enabled = False
        End If
    Next

End Sub

Private Sub Limpa_FrameD1(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataContDe.PromptInclude = False
        DataContDe.Text = ""
        DataContDe.PromptInclude = True
    
        DataContAte.PromptInclude = False
        DataContAte.Text = ""
        DataContAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifCont.ListIndex = -1
        
        ApenasDiasCont.PromptInclude = False
        ApenasDiasCont.Text = ""
        ApenasDiasCont.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasContDe.PromptInclude = False
        EntreDiasContDe.Text = ""
        EntreDiasContDe.PromptInclude = True
        
        EntreQualifContDe.ListIndex = -1
    
        EntreDiasContAte.PromptInclude = False
        EntreDiasContAte.Text = ""
        EntreDiasContAte.PromptInclude = True
    
        EntreQualifContAte.ListIndex = -1
    
    End If

End Sub

Private Sub Limpa_FrameD2(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataPCompDe.PromptInclude = False
        DataPCompDe.Text = ""
        DataPCompDe.PromptInclude = True
    
        DataPCompAte.PromptInclude = False
        DataPCompAte.Text = ""
        DataPCompAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifPComp.ListIndex = -1
        
        ApenasDiasPComp.PromptInclude = False
        ApenasDiasPComp.Text = ""
        ApenasDiasPComp.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasPCompDe.PromptInclude = False
        EntreDiasPCompDe.Text = ""
        EntreDiasPCompDe.PromptInclude = True
        
        EntreQualifPCompDe.ListIndex = -1
    
        EntreDiasPCompAte.PromptInclude = False
        EntreDiasPCompAte.Text = ""
        EntreDiasPCompAte.PromptInclude = True
    
        EntreQualifPCompAte.ListIndex = -1
    
    End If

End Sub

Private Sub Limpa_FrameD3(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataUCompDe.PromptInclude = False
        DataUCompDe.Text = ""
        DataUCompDe.PromptInclude = True
    
        DataUCompAte.PromptInclude = False
        DataUCompAte.Text = ""
        DataUCompAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifUComp.ListIndex = -1
        
        ApenasDiasuComp.PromptInclude = False
        ApenasDiasuComp.Text = ""
        ApenasDiasuComp.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasUCompDe.PromptInclude = False
        EntreDiasUCompDe.Text = ""
        EntreDiasUCompDe.PromptInclude = True
        
        EntreQualifUCompDe.ListIndex = -1
    
        EntreDiasUCompAte.PromptInclude = False
        EntreDiasUCompAte.Text = ""
        EntreDiasUCompAte.PromptInclude = True
    
        EntreQualifUCompAte.ListIndex = -1
    
    End If

End Sub

Public Sub ApenasCont_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_APENAS)
    Call Limpa_FrameD1(FRAMED_APENAS)
End Sub

Public Sub EntreCont_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_ENTRE)
    Call Limpa_FrameD1(FRAMED_ENTRE)
End Sub

Public Sub FaixaDataCont_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call Limpa_FrameD1(FRAMED_FAIXA)
End Sub

Public Sub ApenasPComp_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_APENAS)
    Call Limpa_FrameD2(FRAMED_APENAS)
End Sub

Public Sub EntrePComp_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_ENTRE)
    Call Limpa_FrameD2(FRAMED_ENTRE)
End Sub

Public Sub FaixaDataPComp_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call Limpa_FrameD2(FRAMED_FAIXA)
End Sub

Public Sub ApenasUComp_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_APENAS)
    Call Limpa_FrameD3(FRAMED_APENAS)
End Sub

Public Sub EntreUComp_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_ENTRE)
    Call Limpa_FrameD3(FRAMED_ENTRE)
End Sub

Public Sub FaixaDataUComp_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    Call Limpa_FrameD3(FRAMED_FAIXA)
End Sub

Private Sub Data_Validate(ByVal objDataMask As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se Data está preenchida
    If Len(Trim(objDataMask.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(objDataMask.Text)
        If lErro <> SUCESSO Then gError 200106
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 200106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200107)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 200108
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 200108
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200109)

    End Select
    
End Sub

Public Sub DataContDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataContDe, Cancel)
End Sub

Public Sub DataContAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataContAte, Cancel)
End Sub

Public Sub DataPCompDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataPCompDe, Cancel)
End Sub

Public Sub DataPCompAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataPCompAte, Cancel)
End Sub

Public Sub DataUCompDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataUCompDe, Cancel)
End Sub

Public Sub DataUCompAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataUCompAte, Cancel)
End Sub

Public Sub RespCallCenter_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_RespCallCenter_Validate
    
    'Verifica se algum codigo está selecionado
    If RespCallCenter.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(RespCallCenter.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = RespCallCenter.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 200112
        
        If lErro <> SUCESSO Then gError 200113
        
    End If
    
    Exit Sub
    
Erro_RespCallCenter_Validate:

    Cancel = True

    Select Case gErr
            
        Case 200112
        
        Case 200113 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200114)
    
    End Select
    
    Exit Sub
    
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

'###################################################################
'Inserido por Wagner
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

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200139)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox) As Long

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 200140

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        objCombo.AddItem objCategoriaCliente.sCategoria

    Next
    
    Carrega_ComboCategoriaCliente = SUCESSO
    
    Exit Function

Erro_Carrega_ComboCategoriaCliente:

    Carrega_ComboCategoriaCliente = gErr

    Select Case gErr
    
        Case 200140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200141)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As ComboBox, ByVal objComboItens As ComboBox)

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
        If lErro <> SUCESSO Then gError 200143

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
    
        Case 200143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200144)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 200145
        
        If lErro <> SUCESSO Then gError 200146
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        CategoriaClienteDe.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Enabled = False
        CategoriaClienteAte.Clear
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 200145
         
        Case 200146
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200147)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteItem_Validate(Cancel As Boolean, objCombo As ComboBox)

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Validate

    If Len(objCombo.Text) <> 0 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(objCombo)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 200148
        
        If lErro <> SUCESSO Then gError 200149
    
    End If

    Exit Sub

Erro_CategoriaClienteItem_Validate:

    Cancel = True

    Select Case gErr

        Case 200148
        
        Case 200149
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCombo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200150)

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

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200151)

    End Select

    Exit Sub

End Sub

Public Sub CategoriaClienteAte_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteAte)
End Sub


Public Sub CategoriaClienteDe_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteDe)
End Sub
'####################################################################

Public Sub Selecionado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Selecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Public Sub Selecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Public Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional sValorTexto As String, Optional ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sMes As String
Dim objFilialEmpresa As New AdmFiliais
Dim sTextoAux As String
Dim iIndice As Integer
Dim gcolFiliaisMen As New Collection
Dim gcolEnderecosMen As New Collection
Dim gcolEstatisticaMen As New Collection
Dim gcolClientesMen As New Collection
Dim vValor As Variant
Dim dValorContratos As Double
    
On Error GoTo Erro_Calcula_Mnemonico
   
    'Se não passou a linha o calcula mnemonico vem da tela logo segue a ordenação da tela
    If iLinha = 0 Then
        For Each vValor In gcolFiliais
            gcolFiliaisMen.Add vValor
        Next
        
        For Each vValor In gcolEnderecos
            gcolEnderecosMen.Add vValor
        Next
        
        For Each vValor In gcolEstatisticas
            gcolEstatisticaMen.Add vValor
        Next
        
        For Each vValor In gcolClientes
            gcolClientesMen.Add vValor
        Next
    Else ' Se passou a linha vem do batch logo deve seguir a ordenação do momento do envio
        For Each vValor In gcolFiliaisEnv
            gcolFiliaisMen.Add vValor
        Next
        
        For Each vValor In gcolEnderecosEnv
            gcolEnderecosMen.Add vValor
        Next
        
        For Each vValor In gcolEstatisticasEnv
            gcolEstatisticaMen.Add vValor
        Next
        
        For Each vValor In gcolClientesEnv
            gcolClientesMen.Add vValor
        Next
    End If

    objFilialEmpresa.iCodFilial = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 196985

    sValorTexto = ""
    
    If iLinha <> 0 Then giLinhaAtual = iLinha

    Select Case UCase(objMnemonicoValor.sMnemonico)
                                                      
'        Case UCase(DATA_VENCIMENTO)
'            objMnemonicoValor.colValor.Add DATA_NULA
'            sValorTexto = "  /  /    "
'
'        Case UCase(DATA_VENCIMENTO_REAL)
'            objMnemonicoValor.colValor.Add DATA_NULA
'            sValorTexto = "  /  /    "
'
'        Case UCase(NUMERO_PARCELA)
'            objMnemonicoValor.colValor.Add 0
'            sValorTexto = "0"
'
'        Case UCase(NUMERO_TITULO)
'            objMnemonicoValor.colValor.Add 0
'            sValorTexto = "0"

        Case UCase(RAZAO_CLIENTE)
            objMnemonicoValor.colValor.Add gcolClientesMen.Item(giLinhaAtual).sRazaoSocial
            sValorTexto = gcolClientesMen.Item(giLinhaAtual).sRazaoSocial

'        Case UCase(SALDO_PARCELA)
'            objMnemonicoValor.colValor.Add 0
'            sValorTexto = "0"
'
'        Case UCase(VALOR_PARCELA)
'            objMnemonicoValor.colValor.Add 0
'            sValorTexto = "0"
            
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
            
'        Case UCase(DATA_BAIXA)
'            objMnemonicoValor.colValor.Add DATA_NULA
'            sValorTexto = "  /  /    "
'
        Case UCase(NOME_EMP)
            objMnemonicoValor.colValor.Add gsNomeEmpresa
            sValorTexto = gsNomeEmpresa
            
        Case UCase(CNPJ_EMP)
            objMnemonicoValor.colValor.Add Format(objFilialEmpresa.sCgc, "00\.000\.000\/0000-00; ; ; ")
            sValorTexto = Format(objFilialEmpresa.sCgc, "00\.000\.000\/0000-00; ; ; ")
            
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
            
'        Case UCase(LISTA_NFSPAG)
'
'            sTextoAux = ""
'            objMnemonicoValor.colValor.Add sTextoAux
'            sValorTexto = sTextoAux

        Case UCase(VALOR_CONTRATOS)
        
            dValorContratos = 0
        
            lErro = CF("ContratoCli_Le_Valor", gcolFiliaisMen.Item(giLinhaAtual).lCodCliente, gcolFiliaisMen.Item(giLinhaAtual).iCodFilial, dValorContratos)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objMnemonicoValor.colValor.Add Format(dValorContratos, "STANDARD")
            sValorTexto = Format(dValorContratos, "STANDARD")

        Case UCase(CONTATO_CLIENTE)
            objMnemonicoValor.colValor.Add gcolEnderecosMen.Item(giLinhaAtual).sContato
            sValorTexto = gcolEnderecosMen.Item(giLinhaAtual).sContato
            
        Case Else
            gError 187163

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

Function Verifica_Existencia_Arquivo(ByVal sArquivo As String) As Long

Dim lErro As Long

On Error GoTo Erro_Verifica_Existencia_Arquivo

    If Len(Trim(sArquivo)) > 0 Then

        Open sArquivo For Input As #1
        Close #1
        
    End If

    Verifica_Existencia_Arquivo = SUCESSO

    Exit Function

Erro_Verifica_Existencia_Arquivo:

    Verifica_Existencia_Arquivo = gErr

    Select Case gErr
    
        Case 53
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, sArquivo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189322)

    End Select

    Exit Function

End Function

Public Sub UpDownDataPContAte_DownClick()
    Call UpDownData_DownClick(DataPContAte)
End Sub

Public Sub UpDownDataPContAte_UpClick()
    Call UpDownData_UpClick(DataPContAte)
End Sub

Public Sub UpDownDataPContDe_DownClick()
    Call UpDownData_DownClick(DataPContDe)
End Sub

Public Sub UpDownDataPContDe_UpClick()
    Call UpDownData_UpClick(DataPContDe)
End Sub

Public Sub UpDownDataNCompAte_DownClick()
    Call UpDownData_DownClick(DataNCompAte)
End Sub

Public Sub UpDownDataNCompAte_UpClick()
    Call UpDownData_UpClick(DataNCompAte)
End Sub

Public Sub UpDownDataNCompDe_DownClick()
    Call UpDownData_DownClick(DataNCompDe)
End Sub

Public Sub UpDownDataNCompDe_UpClick()
    Call UpDownData_UpClick(DataNCompDe)
End Sub

Private Sub Limpa_FrameD4(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataPContDe.PromptInclude = False
        DataPContDe.Text = ""
        DataPContDe.PromptInclude = True
    
        DataPContAte.PromptInclude = False
        DataPContAte.Text = ""
        DataPContAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifPCont.ListIndex = -1
        
        ApenasDiasPCont.PromptInclude = False
        ApenasDiasPCont.Text = ""
        ApenasDiasPCont.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasPContDe.PromptInclude = False
        EntreDiasPContDe.Text = ""
        EntreDiasPContDe.PromptInclude = True
        
        EntreQualifPContDe.ListIndex = -1
    
        EntreDiasPContAte.PromptInclude = False
        EntreDiasPContAte.Text = ""
        EntreDiasPContAte.PromptInclude = True
    
        EntreQualifPContAte.ListIndex = -1
    
    End If

End Sub

Private Sub Limpa_FrameD5(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataNCompDe.PromptInclude = False
        DataNCompDe.Text = ""
        DataNCompDe.PromptInclude = True
    
        DataNCompAte.PromptInclude = False
        DataNCompAte.Text = ""
        DataNCompAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        
        ApenasDiasNComp.PromptInclude = False
        ApenasDiasNComp.Text = ""
        ApenasDiasNComp.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasNCompDe.PromptInclude = False
        EntreDiasNCompDe.Text = ""
        EntreDiasNCompDe.PromptInclude = True
    
        EntreDiasNCompAte.PromptInclude = False
        EntreDiasNCompAte.Text = ""
        EntreDiasNCompAte.PromptInclude = True
    
    End If

End Sub

Public Sub ApenasPCont_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_APENAS)
    Call Limpa_FrameD4(FRAMED_APENAS)
End Sub

Public Sub EntrePCont_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_ENTRE)
    Call Limpa_FrameD4(FRAMED_ENTRE)
End Sub

Public Sub FaixaDataPCont_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_FAIXA)
    Call Limpa_FrameD4(FRAMED_FAIXA)
End Sub

Public Sub ApenasNComp_Click()
    Call FrameD_Enabled(FrameD5, FRAMED_APENAS)
    Call Limpa_FrameD5(FRAMED_APENAS)
End Sub

Public Sub EntreNComp_Click()
    Call FrameD_Enabled(FrameD5, FRAMED_ENTRE)
    Call Limpa_FrameD5(FRAMED_ENTRE)
End Sub

Public Sub FaixaDataNComp_Click()
    Call FrameD_Enabled(FrameD5, FRAMED_FAIXA)
    Call Limpa_FrameD5(FRAMED_FAIXA)
End Sub

Public Sub DataNCompDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataNCompDe, Cancel)
End Sub

Public Sub DataNCompAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataNCompAte, Cancel)
End Sub

Public Sub DataPContDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataPContDe, Cancel)
End Sub

Public Sub DataPContAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataPContAte, Cancel)
End Sub

Public Sub TiposCliente_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
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

Function CONTCLIE_Move_Selecao_Memoria(ByVal objCT As Object, ByVal objContatoCliSel As ClassContatoCliSel) As Long
    CONTCLIE_Move_Selecao_Memoria = SUCESSO
End Function

Public Sub IgnoraJaEnviados_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub
