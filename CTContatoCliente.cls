VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTContatoCliente"
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

Dim sMsg As String

Public iAtualizaTela As Integer

Public gobjInfoUsu As Object

Public sOpcaoAnt As String

Public iAlterado As Integer
Public iAlteradoFiltro As Integer
Public iFrameAtual As Integer

Public iStatus_ListIndex_Padrao As Integer

Public objGridClientes As AdmGrid
Public iGrid_CheckLigar_Col As Integer
Public iGrid_FilialCliente_Col As Integer
Public iGrid_Cliente_Col As Integer
Public iGrid_Contato_Col As Integer
Public iGrid_Fone1_Col As Integer
Public iGrid_Fone2_Col As Integer
Public iGrid_CheckLigacaoRealizada_Col As Integer
Public iGrid_Historico_Col As Integer
Public iGrid_ValorCompra_Col As Integer
Public iGrid_DataUCompra_Col As Integer
Public iGrid_DataPCompra_Col As Integer
Public iGrid_DataUContato_Col As Integer

Public gobjContatoCliAnt As ClassContatoCliSel

Public gcolClientes As Collection
Public gcolFiliais As Collection
Public gcolEnderecos As Collection
Public gcolEstatisticas As Collection

Public sCatAnt As String

Const TAB_SELECAO = 1
Const TAB_CLIENTE = 2

Const FRAMED_FAIXA = 0
Const FRAMED_APENAS = 1
Const FRAMED_ENTRE = 2

Dim gbTrazendoDados As Boolean

Dim iConTimer As Integer

'''' API do WIndows
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'--- inicio dos properties get dos controles da tela

Public Property Get FrameTab() As Object
     Set FrameTab = objUserControl.Controls("FrameTab")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
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

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get BotaoGravarContato() As Object
     Set BotaoGravarContato = objUserControl.Controls("BotaoGravarContato")
End Property

Public Property Get BotaoContatos() As Object
     Set BotaoContatos = objUserControl.Controls("BotaoContatos")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get Assunto() As Object
     Set Assunto = objUserControl.Controls("Assunto")
End Property

Public Property Get lblCliente() As Object
     Set lblCliente = objUserControl.Controls("lblCliente")
End Property

Public Property Get LabelAssunto() As Object
     Set LabelAssunto = objUserControl.Controls("LabelAssunto")
End Property

Public Property Get ValorTotalCompras() As Object
     Set ValorTotalCompras = objUserControl.Controls("ValorTotalCompras")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get OperDDD() As Object
     Set OperDDD = objUserControl.Controls("OperDDD")
End Property

Public Property Get SenhaTel() As Object
     Set SenhaTel = objUserControl.Controls("SenhaTel")
End Property

Public Property Get BotaoLigar() As Object
     Set BotaoLigar = objUserControl.Controls("BotaoLigar")
End Property

Public Property Get BotaoDesligar() As Object
     Set BotaoDesligar = objUserControl.Controls("BotaoDesligar")
End Property

Public Property Get Contato() As Object
     Set Contato = objUserControl.Controls("Contato")
End Property

Public Property Get PossuiSenha() As Object
     Set PossuiSenha = objUserControl.Controls("PossuiSenha")
End Property

Public Property Get DigDiscExt() As Object
     Set DigDiscExt = objUserControl.Controls("DigDiscExt")
End Property

Public Property Get UF() As Object
     Set UF = objUserControl.Controls("UF")
End Property

Public Property Get LabelContato() As Object
     Set LabelContato = objUserControl.Controls("LabelContato")
End Property

Public Property Get Telefone() As Object
     Set Telefone = objUserControl.Controls("Telefone")
End Property

Public Property Get BotaoMarcar() As Object
     Set BotaoMarcar = objUserControl.Controls("BotaoMarcar")
End Property

Public Property Get BotaoDesmarcar() As Object
     Set BotaoDesmarcar = objUserControl.Controls("BotaoDesmarcar")
End Property

Public Property Get BotaoHistorico() As Object
     Set BotaoHistorico = objUserControl.Controls("BotaoHistorico")
End Property

Public Property Get BotaoGravarTela() As Object
     Set BotaoGravarTela = objUserControl.Controls("BotaoGravarTela")
End Property

Public Property Get BotaoCliente() As Object
     Set BotaoCliente = objUserControl.Controls("BotaoCliente")
End Property

Public Property Get HistoricoGrid() As Object
     Set HistoricoGrid = objUserControl.Controls("HistoricoGrid")
End Property

Public Property Get Fone2Grid() As Object
     Set Fone2Grid = objUserControl.Controls("Fone2Grid")
End Property

Public Property Get CheckLigacaoRealizada() As Object
     Set CheckLigacaoRealizada = objUserControl.Controls("CheckLigacaoRealizada")
End Property

Public Property Get CheckLigar() As Object
     Set CheckLigar = objUserControl.Controls("CheckLigar")
End Property

Public Property Get Fone1Grid() As Object
     Set Fone1Grid = objUserControl.Controls("Fone1Grid")
End Property

Public Property Get ClienteGrid() As Object
     Set ClienteGrid = objUserControl.Controls("ClienteGrid")
End Property

Public Property Get FilialClienteGrid() As Object
     Set FilialClienteGrid = objUserControl.Controls("FilialClienteGrid")
End Property

Public Property Get ContatoGrid() As Object
     Set ContatoGrid = objUserControl.Controls("ContatoGrid")
End Property

Public Property Get GridClientes() As Object
     Set GridClientes = objUserControl.Controls("GridClientes")
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

Public Property Get RespCallCenter() As Object
     Set RespCallCenter = objUserControl.Controls("RespCallCenter")
End Property

Public Property Get FrameD3() As Object
     Set FrameD3 = objUserControl.Controls("FrameD3")
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
     Set ApenasDiasuComp = objUserControl.Controls("ApenasDiasUComp")
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

Public Property Get EntreUComp() As Object
     Set EntreUComp = objUserControl.Controls("EntreUComp")
End Property

Public Property Get ApenasUComp() As Object
     Set ApenasUComp = objUserControl.Controls("ApenasUComp")
End Property

Public Property Get FaixaDataUComp() As Object
     Set FaixaDataUComp = objUserControl.Controls("FaixaDataUComp")
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

Public Property Get EntreNComp() As Object
     Set EntreNComp = objUserControl.Controls("EntreNComp")
End Property

Public Property Get ApenasNComp() As Object
     Set ApenasNComp = objUserControl.Controls("ApenasNComp")
End Property

Public Property Get FaixaDataNComp() As Object
     Set FaixaDataNComp = objUserControl.Controls("FaixaDataNComp")
End Property

Public Property Get EntreQualifPContDe() As Object
     Set EntreQualifPContDe = objUserControl.Controls("EntreQualifPContDe")
End Property

Public Property Get EntreQualifPContAte() As Object
     Set EntreQualifPContAte = objUserControl.Controls("EntreQualifPContAte")
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

Public Property Get EntrePCont() As Object
     Set EntrePCont = objUserControl.Controls("EntrePCont")
End Property

Public Property Get ApenasPCont() As Object
     Set ApenasPCont = objUserControl.Controls("ApenasPCont")
End Property

Public Property Get FaixaDataPCont() As Object
     Set FaixaDataPCont = objUserControl.Controls("FaixaDataPCont")
End Property

Public Property Get FrameD2() As Object
     Set FrameD2 = objUserControl.Controls("FrameD2")
End Property

Public Property Get EntreQualifPCompDe() As Object
     Set EntreQualifPCompDe = objUserControl.Controls("EntreQualifPCompDe")
End Property

Public Property Get EntreQualifPCompAte() As Object
     Set EntreQualifPCompAte = objUserControl.Controls("EntreQualifPCompAte")
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

Public Property Get EntrePComp() As Object
     Set EntrePComp = objUserControl.Controls("EntrePComp")
End Property

Public Property Get ApenasPComp() As Object
     Set ApenasPComp = objUserControl.Controls("ApenasPComp")
End Property

Public Property Get FaixaDataPComp() As Object
     Set FaixaDataPComp = objUserControl.Controls("FaixaDataPComp")
End Property

Public Property Get FrameD1() As Object
     Set FrameD1 = objUserControl.Controls("FrameD1")
End Property

Public Property Get FrameD5() As Object
     Set FrameD5 = objUserControl.Controls("FrameD5")
End Property

Public Property Get FrameD4() As Object
     Set FrameD4 = objUserControl.Controls("FrameD4")
End Property

Public Property Get EntreQualifContAte() As Object
     Set EntreQualifContAte = objUserControl.Controls("EntreQualifContAte")
End Property

Public Property Get EntreQualifContDe() As Object
     Set EntreQualifContDe = objUserControl.Controls("EntreQualifContDe")
End Property

Public Property Get EntreDiasContDe() As Object
     Set EntreDiasContDe = objUserControl.Controls("EntreDiasContDe")
End Property

Public Property Get EntreDiasContAte() As Object
     Set EntreDiasContAte = objUserControl.Controls("EntreDiasContAte")
End Property

Public Property Get ApenasQualifCont() As Object
     Set ApenasQualifCont = objUserControl.Controls("ApenasQualifCont")
End Property

Public Property Get ApenasDiasCont() As Object
     Set ApenasDiasCont = objUserControl.Controls("ApenasDiasCont")
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

Public Property Get FaixaDataCont() As Object
     Set FaixaDataCont = objUserControl.Controls("FaixaDataCont")
End Property

Public Property Get ApenasCont() As Object
     Set ApenasCont = objUserControl.Controls("ApenasCont")
End Property

Public Property Get EntreCont() As Object
     Set EntreCont = objUserControl.Controls("EntreCont")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get OpcaoPadrao() As Object
     Set OpcaoPadrao = objUserControl.Controls("OpcaoPadrao")
End Property

Public Property Get OpcoesTela() As Object
     Set OpcoesTela = objUserControl.Controls("OpcoesTela")
End Property

Public Property Get LabelOpcao() As Object
     Set LabelOpcao = objUserControl.Controls("LabelOpcao")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get ComDiscar() As Object
     Set ComDiscar = objUserControl.Controls("ComDiscar")
End Property

Public Property Get Timer1() As Object
     Set Timer1 = objUserControl.Controls("Timer1")
End Property

Public Property Get Timer2() As Object
     Set Timer2 = objUserControl.Controls("Timer2")
End Property

Public Property Get DataProx() As Object
     Set DataProx = objUserControl.Controls("DataProx")
End Property

Public Property Get TiposCliente() As Object
     Set TiposCliente = objUserControl.Controls("TiposCliente")
End Property

Public Property Get SoAtivos() As Object
     Set SoAtivos = objUserControl.Controls("SoAtivos")
End Property

'--- fim dos properties get dos controles da tela

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Contato com Clientes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContatoCliente"

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
    
    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objGridClientes = Nothing
    Set gcolClientes = Nothing
    Set gcolFiliais = Nothing
    Set gcolEnderecos = Nothing
    Set gcolEstatisticas = Nothing

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
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    gbTrazendoDados = False

    Set objGridClientes = New AdmGrid
    
    Set gobjContatoCliAnt = New ClassContatoCliSel
    
    Set gcolClientes = New Collection
    Set gcolFiliais = New Collection
    Set gcolEnderecos = New Collection
    Set gcolEstatisticas = New Collection
    
    lErro = Inicializa_GridClientes(objGridClientes)
    If lErro <> SUCESSO Then gError 200036

    '#################################################
    'Inserido por Wagner
    Call Carrega_ComboCategoriaCliente(CategoriaCliente)
    '#################################################

    Call Carrega_Usuarios
    
    Call Carrega_Status(Status)
    
    '#####################################
    'Inserido por Wagner
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    '#####################################

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
    
    sDigDiscExt = String(128, 0)
    
    Call GetPrivateProfileString("Geral", "DigDiscExt", "", sDigDiscExt, 128, "ADM100.INI")

    sDigDiscExt = Replace(sDigDiscExt, Chr(0), "")

    DigDiscExt.PromptInclude = False
    DigDiscExt.Text = sDigDiscExt
    DigDiscExt.PromptInclude = True
    
    BotaoDesligar.Enabled = False
    
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
    
        Case 200036, 200037, 200224
            'erros tratados nas rotinas chamadas

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

    If Len(Trim(RespCallCenter.Text)) = 0 Then gError 200040
    
    objContatoCliSel.sRespCallCenter = RespCallCenter.Text
    
    If SoAtivos.Value = vbChecked Then
        objContatoCliSel.iSoAtivos = MARCADO
    Else
        objContatoCliSel.iSoAtivos = DESMARCADO
    End If

    If FaixaDataCont.Value Then
        objContatoCliSel.dtDataContDe = StrParaDate(DataContDe.Text)
        objContatoCliSel.dtDataContAte = StrParaDate(DataContAte.Text)
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
    
    lErro = CF2(Me, "CONTCLI_Move_Selecao_Memoria", objContatoCliSel)
    If lErro <> SUCESSO Then gError 200986
   
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
            
        Case 200986

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200055)

    End Select

    Exit Function

End Function

Function CONTCLI_Move_Selecao_Memoria(ByVal objCT As Object, ByVal objContatoCliSel As ClassContatoCliSel) As Long
    CONTCLI_Move_Selecao_Memoria = SUCESSO
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

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("ContatoCli_Selecao_Le", objContatoCliSel, colFiliais, colEnderecos, colClienteEst)
    If lErro <> SUCESSO Then gError 200061
    
    If colFiliais.Count = 0 Then gError 200062

    lErro = Preenche_GridCliente(colFiliais, colEnderecos, colClienteEst)
    If lErro <> SUCESSO Then gError 200063
    
    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 200061, 200063
        
        Case 200062
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_COBRANCA_SEM_CLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200064)

    End Select

    Exit Function

End Function

Function Preenche_GridCliente(ByVal colFiliais As Collection, ByVal colEnderecos As Collection, ByVal colEstatisticas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilial As ClassFilialCliente
Dim objEnderecos As ClassEndereco
Dim colCliente As New Collection
Dim objCliente As ClassCliente
Dim bAchou As Boolean
Dim objFilialContatoData As New ClassFilialContatoData
Dim objClieEst As ClassFilialClienteEst
Dim iPos As Integer

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
        Set objEnderecos = colEnderecos(iIndice)
        
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
        
        objFilialContatoData.lCliente = objFilial.lCodCliente
        objFilialContatoData.iFilial = objFilial.iCodFilial
        objFilialContatoData.dtData = gdtDataAtual
        
        'Busca dados de Ligação, histórico, Ligações efetuadas naquela data
        lErro = CF("FilialContatoData_Le", objFilialContatoData)
        If lErro <> SUCESSO Then gError 200066

        GridClientes.TextMatrix(iIndice, iGrid_CheckLigacaoRealizada_Col) = objFilialContatoData.iLigacaoEfetuada
        GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = objFilialContatoData.iLigar
        GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        GridClientes.TextMatrix(iIndice, iGrid_FilialCliente_Col) = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
        GridClientes.TextMatrix(iIndice, iGrid_Historico_Col) = objFilialContatoData.sHistorico
'        GridClientes.TextMatrix(iIndice, iGrid_Fone1_Col) = objEnderecos.sTelefone1
'        GridClientes.TextMatrix(iIndice, iGrid_Fone2_Col) = objEnderecos.sTelefone2
'        GridClientes.TextMatrix(iIndice, iGrid_Contato_Col) = objEnderecos.sContato
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
        gcolEnderecos.Add objEnderecos
        gcolEstatisticas.Add objClieEst
        gcolFiliais.Add objFilial
        
    Next
           
    objGridClientes.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridClientes)

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
    objGrid.colColuna.Add ("Ligar")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Vlr.Compra")
    objGrid.colColuna.Add ("Últ.Contato")
    objGrid.colColuna.Add ("Pri.Compra")
    objGrid.colColuna.Add ("Últ.Compra")
    objGrid.colColuna.Add ("Lig.Efet.")
    objGrid.colColuna.Add ("Histórico")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (CheckLigar.Name)
    objGrid.colCampo.Add (ClienteGrid.Name)
    objGrid.colCampo.Add (FilialClienteGrid.Name)
    objGrid.colCampo.Add (ValorTotalCompras.Name)
    objGrid.colCampo.Add (DataUContato.Name)
    objGrid.colCampo.Add (DataPCompra.Name)
    objGrid.colCampo.Add (DataUCompra.Name)
    objGrid.colCampo.Add (CheckLigacaoRealizada.Name)
    objGrid.colCampo.Add (HistoricoGrid.Name)

    'Colunas do Grid
    iGrid_CheckLigar_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_FilialCliente_Col = 3
    iGrid_ValorCompra_Col = 4
    iGrid_DataUContato_Col = 5
    iGrid_DataPCompra_Col = 6
    iGrid_DataUCompra_Col = 7
    iGrid_CheckLigacaoRealizada_Col = 8
    iGrid_Historico_Col = 9

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

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridClientes)
    
    If Not gbTrazendoDados Then
        If objGridClientes.iLinhaAntiga <> objGridClientes.objGrid.Row Then
            Call Trata_Contato
            objGridClientes.iLinhaAntiga = objGridClientes.objGrid.Row
        End If
    End If
    
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

                Case iGrid_CheckLigar_Col

                    lErro = Saida_Celula_Ligar(objGridInt)
                    If lErro <> SUCESSO Then gError 200071

                Case iGrid_CheckLigacaoRealizada_Col

                    lErro = Saida_Celula_LigacaoReal(objGridInt)
                    If lErro <> SUCESSO Then gError 200072
                    
                Case iGrid_Historico_Col

                    lErro = Saida_Celula_Historico(objGridInt)
                    If lErro <> SUCESSO Then gError 200073

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

Private Function Saida_Celula_Ligar(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Ligar do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Ligar

    Set objGridInt.objControle = CheckLigar

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200076

    Saida_Celula_Ligar = SUCESSO

    Exit Function

Erro_Saida_Celula_Ligar:

    Saida_Celula_Ligar = gErr

    Select Case gErr
        
        Case 200076
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200077)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_LigacaoReal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula LigacaoReal do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_LigacaoReal

    Set objGridInt.objControle = CheckLigacaoRealizada

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200078

    Saida_Celula_LigacaoReal = SUCESSO

    Exit Function

Erro_Saida_Celula_LigacaoReal:

    Saida_Celula_LigacaoReal = gErr

    Select Case gErr
        
        Case 200078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200079)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Historico do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Historico

    Set objGridInt.objControle = HistoricoGrid

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200080

    Saida_Celula_Historico = SUCESSO

    Exit Function

Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = gErr

    Select Case gErr
        
        Case 200080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200081)

    End Select

    Exit Function

End Function

Public Sub BotaoDesligar_Click()
'Desconecta a porta do modem

On Error GoTo Erro_BotaoLigar_Click

    If ComDiscar.PortOpen Then
       ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If

    Exit Sub

Erro_BotaoLigar_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200082)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoLigar_Click()
'Faz a ligação telefonica usando o modem

Dim sCOM As String
Dim sTelefone As String
Dim sTelefoneCompl As String
Dim iCont As Integer
Dim sOperDDDViaPABX As String

On Error GoTo Erro_BotaoLigar_Click

    iConTimer = 0
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer1.Interval = 100
    Timer2.Interval = 1

    If GridClientes.Row = 0 Then gError 200084
    If ComDiscar.PortOpen Then gError 200085
    If Len(Trim(Telefone.Caption)) = 0 Then gError 200086
    
    sCOM = String(128, 0)
    sOperDDDViaPABX = String(128, 0)
    Call GetPrivateProfileString("Geral", "modem", "3", sCOM, 128, "ADM100.INI")
    Call GetPrivateProfileString("Geral", "OperDDDViaPABX", "0", sOperDDDViaPABX, 128, "ADM100.INI")
    sOperDDDViaPABX = Replace(sOperDDDViaPABX, Chr(0), "")
    sCOM = Replace(sCOM, Chr(0), "")
       
    ComDiscar.CommPort = CInt(sCOM)
    ComDiscar.Settings = "9600,N,8,1"
    ComDiscar.PortOpen = True
    
    BotaoDesligar.Enabled = True
    DoEvents
    
    sTelefone = Telefone.Caption

    sTelefone = Replace(sTelefone, "(", "")
    sTelefone = Replace(sTelefone, ")", "")
    sTelefone = Replace(sTelefone, "-", "")
    sTelefone = Replace(sTelefone, " ", "")
    
    'Se tem Senha
    If PossuiSenha.Value = vbChecked Then
        sTelefoneCompl = sTelefoneCompl & "#" & Trim(SenhaTel.Text) & ","
    End If
    
    'Se tem dígito para ligações externas
    If Len(Trim(DigDiscExt.Text)) > 0 Then
        sTelefoneCompl = sTelefoneCompl & Trim(DigDiscExt.Text) & ","
    End If
    
    'Se é um DDD
    If Len(sTelefone) > 9 Then
    
        'Coloca o 0 na frente
        sTelefoneCompl = sTelefoneCompl & "0"
        
        'Se a operadora de DDD não é obtida do PABX
        If StrParaInt(sOperDDDViaPABX) = DESMARCADO Then
        
            'Coloca o código da operadora de DDD
            sTelefoneCompl = sTelefoneCompl & OperDDD.Text
        End If
    End If
    
    sTelefone = sTelefoneCompl & sTelefone
    
    ComDiscar.Output = "ATDT" & sTelefone & vbCr
    
    'Call BotaoHistorico_Click
    
    Exit Sub

Erro_BotaoLigar_Click:

    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If

    Select Case gErr
    
        Case 200084
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 200085
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_SINAL_LINHA", gErr)
             
        Case 200086
             Call Rotina_Erro(vbOKOnly, "ERRO_TELEFONE_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200087)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoGravarTela_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialContatoData As ClassFilialContatoData
Dim colFilialContatoData As New Collection

On Error GoTo Erro_BotaoGravarTela_Click

    GL_objMDIForm.MousePointer = vbHourglass

    For iIndice = 1 To objGridClientes.iLinhasExistentes
    
        Set objFilialContatoData = New ClassFilialContatoData
    
        objFilialContatoData.dtData = gdtDataAtual
        objFilialContatoData.iLigacaoEfetuada = StrParaInt(GridClientes.TextMatrix(iIndice, iGrid_CheckLigacaoRealizada_Col))
        objFilialContatoData.iLigar = StrParaInt(GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col))
        objFilialContatoData.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col))
        objFilialContatoData.iFilial = Codigo_Extrai(GridClientes.TextMatrix(iIndice, iGrid_FilialCliente_Col))
        objFilialContatoData.sHistorico = GridClientes.TextMatrix(iIndice, iGrid_Historico_Col)
        objFilialContatoData.sCodUsuario = RespCallCenter.Text
    
        colFilialContatoData.Add objFilialContatoData
    
    Next
    
    lErro = CF("FilialContatoData_Grava", colFilialContatoData)
    If lErro <> SUCESSO Then gError 200088
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_GRAVACAO_COM_SUCESSO")
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGravarTela_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 200088

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200089)

    End Select

    Exit Sub
    
End Sub

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
    
    SoAtivos.Value = vbUnchecked
    
    sOpcaoAnt = ""
    sCatAnt = ""
    
    Status.ListIndex = iStatus_ListIndex_Padrao
    
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

Private Sub UpDownData_DownClick(objDataMask As Object)

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

Private Sub UpDownData_UpClick(objDataMask As Object)

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
            GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = MARCADO
        Else
            GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = DESMARCADO
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

Private Sub Data_Validate(objDataMask As Object, Cancel As Boolean)

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

Public Sub BotaoHistorico_Click()

Dim objHistCobr As New ClassHistoricoCobrSelCli

On Error GoTo Erro_BotaoHistorico_Click

    If GridClientes.Row = 0 Then gError 200110

    objHistCobr.iFilial = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objHistCobr.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    objHistCobr.iContato = Codigo_Extrai(Contato.Text)

    Call Chama_Tela("HistoricoCliente", objHistCobr)

    Exit Sub

Erro_BotaoHistorico_Click:

    Select Case gErr
    
        Case 200110
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200111)

    End Select

    Exit Sub
    
End Sub

Public Sub RespCallCenter_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_RespCallCenter_Validate
    
    'Verifica se algum codigo está selecionado
    'If RespCallCenter.ListIndex = -1 Then Exit Sub
    
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

Public Function Trata_Contato() As Long

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim sOperDDD As String
Dim sFilialMaquina As String
Dim lCliente As Long
Dim iFilial As Integer
Dim iIndice As Integer

On Error GoTo Erro_Trata_Contato

    lblCliente.Caption = ""

    If GridClientes.Row <> 0 And objGridClientes.iLinhasExistentes <> 0 Then
    
        Set objEndereco = gcolEnderecos.Item(GridClientes.Row)
    
        lblCliente.Caption = GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col)

        lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
        iFilial = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))

        'Guarda no objClienteContatos, o código do cliente e da
        objClienteContatos.lCliente = lCliente
        objClienteContatos.iFilialCliente = iFilial
        
        'Carrega a combo de contatos
        lErro = CF("Carrega_ClienteContatos", Contato, objClienteContatos, objEndereco.sTelefone1, objEndereco.sContato)
        If lErro <> SUCESSO And lErro <> 102622 Then gError 200115
                
        'Se selecionou o contato padrão =>
        If Len(Trim(Contato.Text)) > 0 Then
        
            'traz o telefone do contato
            Call Contato_Click
        
        Else
        
            'Limpa o campo telefone
            Telefone.Caption = ""
            
        End If

        
'        If lCliente <> 0 And iFilial <> 0 Then
'
'            objFilialCliente.lCodCliente = lCliente
'            objFilialCliente.iCodFilial = iFilial
'
'            'Lê os dados da Filial Cliente
'            lErro = CF("FilialCliente_Le", objFilialCliente)
'            If lErro <> SUCESSO And lErro <> 12567 Then gError 200116
'
'            objEndereco.lCodigo = objFilialCliente.lEndereco
'
'            lErro = CF("Endereco_Le", objEndereco)
'            If lErro <> SUCESSO And lErro <> 12309 Then gError 200117
            
            UF.Caption = objEndereco.sSiglaEstado
            
            sFilialMaquina = String(128, 0)
            
            Call GetPrivateProfileString("Geral", "FilialMaquina", "1", sFilialMaquina, 128, "ADM100.INI")
        
            sFilialMaquina = Replace(sFilialMaquina, Chr(0), "")
            
            lErro = CF("OperadorasDDD_Le", StrParaInt(sFilialMaquina), UF.Caption, sOperDDD)
            If lErro <> SUCESSO Then gError 200118
            
            OperDDD.PromptInclude = False
            OperDDD.Text = sOperDDD
            OperDDD.PromptInclude = True
            
'        Else
'
'            UF.Caption = ""
'
'            OperDDD.PromptInclude = False
'            OperDDD.Text = ""
'            OperDDD.PromptInclude = True
'
'        End If
'
    End If
    
    Trata_Contato = SUCESSO

    Exit Function

Erro_Trata_Contato:

    Trata_Contato = gErr

    Select Case gErr

        Case 200115 To 200118
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200119)

    End Select

End Function

Public Sub Contato_Click()

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Contato_Click

    'Se o campo contato não foi preenchido => sai da função
    If Contato.ListIndex = -1 Then Exit Sub
    
    If GridClientes.Row = 0 Then gError 200120

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    objClienteContatos.iCodigo = Codigo_Extrai(Contato.Text)

    'Lê o contato no BD
    lErro = CF("ClienteContatos_Le", objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102653 Then gError 200121
    
    'Se não encontrou o contato => erro
    If lErro = 102653 Then gError 200122
    
    'Exibe o telefone cadastrado para o contato selecionado
    Telefone.Caption = objClienteContatos.sTelefone
    
    Exit Sub
    
Erro_Contato_Click:

    Select Case gErr

        Case 200120
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 200121
        
        Case 200122
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Contato.Text, objClienteContatos.lCliente, objClienteContatos.iFilialCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200123)

    End Select
    

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Public Sub Contato_Validate(Cancel As Boolean)
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos
Dim iCodigo As Integer

On Error GoTo Erro_Contato_Validate

    'Se o contato foi preenchido
    If Len(Trim(Contato.Text)) > 0 Then
    
        If GridClientes.Row = 0 Then gError 200124
    
        'Guarda o código do cliente e da filial no obj
        objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
        objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    
        'Se o contato foi selecionado na própria combo => sai da função
        If Contato.Text = Contato.List(Contato.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Contato, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 200125
    
        'Se não encontrou o contato na combo, mas retornou um código
        If lErro = 6730 Then

            objClienteContatos.iCodigo = iCodigo
            
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 102653 Then gError 200126
            
            'Se não encontrou o contato
            If lErro = 102653 Then gError 200127
            
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Caption = objClienteContatos.sTelefone
        
        End If
        
        'Se foi digitado o nome do contato
        'e esse nome não foi encontrado na combo => erro
        If lErro = 6731 Then

            objClienteContatos.sContato = Contato.Text
        
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le_Nome", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 178440 Then gError 200128
            
            'Se não encontrou o contato
            If lErro = 178440 Then gError 200129
        
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Caption = objClienteContatos.sTelefone
        
        End If
    
    'Senão
    Else
    
        'Limpa o campo telefone
        Telefone.Caption = ""
    
    End If
    
    Exit Sub

Erro_Contato_Validate:

    Cancel = True

    Select Case gErr

        Case 200124
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 200125, 200126, 200128
            
        Case 200127, 200129
            
            'Verifica se o usuário deseja criar um novo contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTECONTATO", Contato.Text, GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col), GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

            'Se o usuário respondeu sim
            If vbMsgRes = vbYes Then
                'Chama a tela para cadastro de contatos
                Call Chama_Tela("ClienteContatos", objClienteContatos)
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200130)

    End Select

    Exit Sub

End Sub

Public Sub LabelContato_Click()

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_LabelContato_Click

    If GridClientes.Row = 0 Then gError 200131

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    
    Call Chama_Tela("ClienteContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelContato_Click:

    Select Case gErr

        Case 200131
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 200132)

    End Select

    Exit Sub
    
End Sub

Public Sub HistoricoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub HistoricoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Public Sub HistoricoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Public Sub HistoricoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = HistoricoGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CheckLigar_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CheckLigar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Public Sub CheckLigar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Public Sub CheckLigar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = CheckLigar
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CheckLigacaoRealizada_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CheckLigacaoRealizada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Public Sub CheckLigacaoRealizada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Public Sub CheckLigacaoRealizada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = CheckLigacaoRealizada
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoCliente_Click()

Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoCliente_Click

    If GridClientes.Row = 0 Then gError 200133

    objCliente.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

    Call Chama_Tela("Clientes", objCliente)

    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr
    
        Case 200133
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200134)

    End Select

    Exit Sub
    
End Sub

Public Sub Timer1_Timer()

On Error GoTo Erro_Timer1_Timer

    If sMsg <> "" Then

        If InStr(1, sMsg, "NO DIALTONE") <> 0 Then gError 200135
'        If InStr(1, sMsg, "NO ANSWER") <> 0 Then gError 200135
        If InStr(1, sMsg, "BUSY") <> 0 Then gError 200136
'        If InStr(1, sMsg, "NO CARRIER") <> 0 Then gError 200135
    
        If InStr(1, sMsg, "RING BACK") <> 0 Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_TEL_TOCANDO")
            Timer1.Interval = 0
            Timer2.Interval = 0
            sMsg = ""
        End If
        
    End If
    
    iConTimer = iConTimer + 1
    
    If iConTimer > 50 Then
        'gError 200135
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_TEL_TOCANDO")
        Timer1.Interval = 0
        Timer2.Interval = 0
        sMsg = ""
    End If

    Exit Sub

Erro_Timer1_Timer:

    sMsg = ""
    Timer1.Interval = 0
    Timer2.Interval = 0
    
    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
    End If

    Select Case gErr

        Case 200135
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_SINAL_LINHA", gErr)
             
        Case 200136
             Call Rotina_Erro(vbOKOnly, "ERRO_TEL_OCUPADO", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200137)

    End Select

    Exit Sub

End Sub

Public Sub Timer2_Timer()

On Error GoTo Erro_Timer2_Timer

    If ComDiscar.PortOpen Then

        sMsg = sMsg + ComDiscar.Input
        
    End If

    Exit Sub

Erro_Timer2_Timer:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200138)

    End Select

    Exit Sub

End Sub

Public Sub PossuiSenha_Click()

    If PossuiSenha.Value = vbChecked Then
        SenhaTel.Enabled = True
    Else
        SenhaTel.Enabled = False
        SenhaTel.Text = ""
    End If

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
        
    sCatAnt = CategoriaCliente.Text
    
    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200139)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox)

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
    
    Exit Sub

Erro_Carrega_ComboCategoriaCliente:

    Select Case gErr
    
        Case 200140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200141)

    End Select

    Exit Sub

End Sub

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

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSRELACCLI, objComboBox)
    If lErro <> SUCESSO Then gError 200152

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iStatus_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 200152

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200153)

    End Select

    Exit Function

End Function

Public Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoContatos_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoContatos_Click

    If GridClientes.Row = 0 Then gError 200154
        
    colSelecao.Add LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, Nothing, "Cliente = ? ")

    Exit Sub

Erro_BotaoContatos_Click:

    Select Case gErr
    
        Case 200154
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200155)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravarContato_Click()

Dim lErro As Long
Dim objRelacCli As New ClassRelacClientes

On Error GoTo Erro_BotaoGravarContato_Click

    If GridClientes.Row = 0 Then gError 200156

    objRelacCli.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    objRelacCli.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objRelacCli.dtData = gdtDataAtual
    objRelacCli.dtDataPrevReceb = DATA_NULA
    objRelacCli.dtDataProxCobr = StrParaDate(DataProx.Text)
    objRelacCli.iFilialEmpresa = giFilialEmpresa
    'Guarda no obj a primeira parte do assunto
    objRelacCli.sAssunto1 = left(Assunto.Text, STRING_BUFFER_MAX_TEXTO - 1)
    'Guarda no obj a segunda parte do assunto
    objRelacCli.sAssunto2 = Mid(Assunto.Text, STRING_BUFFER_MAX_TEXTO)
    objRelacCli.lNumIntParcRec = 0
    objRelacCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    objRelacCli.dtHora = Time
    objRelacCli.iContato = Codigo_Extrai(Contato.Text)
    objRelacCli.lTipo = TIPO_RELACIONAMENTO_CONTATO
    
    If Status.ListIndex <> -1 Then
        objRelacCli.iStatusCG = Status.ItemData(Status.ListIndex)
    Else
        objRelacCli.iStatusCG = 0
    End If
    
    lErro = CF("RelacionamentoClientes_Grava", objRelacCli, True, gsUsuario)
    If lErro <> SUCESSO Then gError 200157
    
    Assunto.Text = ""
    Status.ListIndex = iStatus_ListIndex_Padrao
    
    DataProx.PromptInclude = False
    DataProx.Text = ""
    DataProx.PromptInclude = True
    
    Exit Sub
    
Erro_BotaoGravarContato_Click:

    Select Case gErr
    
        Case 200156
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 200157
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200158)

    End Select
    
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

Public Sub UpDownDataProx_DownClick()
    Call UpDownData_DownClick(DataProx)
End Sub

Public Sub UpDownDataProx_UpClick()
    Call UpDownData_UpClick(DataProx)
End Sub

Public Sub DataProx_Validate(Cancel As Boolean)
    Call Data_Validate(DataProx, Cancel)
End Sub

Public Sub DataProx_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TiposCliente_Click()
    iAlteradoFiltro = REGISTRO_ALTERADO
End Sub
