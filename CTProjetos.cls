VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

Public gobjInfoUsu As Object 'Inserido por Wagner
Dim gobjAnexos As ClassAnexos

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim gobjTelaEscopo As ClassTelaEscopo
Dim gobjTelaCamposCust As ClassTelaDadosCust

Dim gobjProjeto As ClassProjetos

Dim iFrameAtual As Integer
Public iAlterado As Integer

Const INICIO_PROJETO = "Inicio_Projeto"
Const INICIO_ETAPA = "Inicio_Etapa"
Const PERC_COMPL_PRJ = "Perc_Compl_PRJ"
Const PERC_COMPL_ETAPA = "Perc_Compl_Etapa"

'--- inicio dos properties get dos controles da tela

Public Property Get RespUsu() As Object
     Set RespUsu = objUserControl.Controls("RespUsu")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get TEXTO() As Object
     Set TEXTO = objUserControl.Controls("Texto")
End Property

Public Property Get BotaoDadosCustDel() As Object
     Set BotaoDadosCustDel = objUserControl.Controls("BotaoDadosCustDel")
End Property

Public Property Get Controles() As Object
     Set Controles = objUserControl.Controls("Controles")
End Property

Public Property Get BotaoDadosCustNovo() As Object
     Set BotaoDadosCustNovo = objUserControl.Controls("BotaoDadosCustNovo")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get EscExclusoes() As Object
     Set EscExclusoes = objUserControl.Controls("EscExclusoes")
End Property

Public Property Get EscPremissas() As Object
     Set EscPremissas = objUserControl.Controls("EscPremissas")
End Property

Public Property Get EscRestricoes() As Object
     Set EscRestricoes = objUserControl.Controls("EscRestricoes")
End Property

Public Property Get EscFatores() As Object
     Set EscFatores = objUserControl.Controls("EscFatores")
End Property

Public Property Get EscExpectativa() As Object
     Set EscExpectativa = objUserControl.Controls("EscExpectativa")
End Property

Public Property Get EscDescricao() As Object
     Set EscDescricao = objUserControl.Controls("EscDescricao")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get DataFimCalc() As Object
     Set DataFimCalc = objUserControl.Controls("DataFimCalc")
End Property

Public Property Get DataInicioCalc() As Object
     Set DataInicioCalc = objUserControl.Controls("DataInicioCalc")
End Property

Public Property Get Duracao() As Object
     Set Duracao = objUserControl.Controls("Duracao")
End Property

Public Property Get Intervalo() As Object
     Set Intervalo = objUserControl.Controls("Intervalo")
End Property

Public Property Get UpDownDataInicio() As Object
     Set UpDownDataInicio = objUserControl.Controls("UpDownDataInicio")
End Property

Public Property Get DataInicioInf() As Object
     Set DataInicioInf = objUserControl.Controls("DataInicioInf")
End Property

Public Property Get UpDownDataFim() As Object
     Set UpDownDataFim = objUserControl.Controls("UpDownDataFim")
End Property

Public Property Get DataFimInf() As Object
     Set DataFimInf = objUserControl.Controls("DataFimInf")
End Property

Public Property Get DataFimRealCalc() As Object
     Set DataFimRealCalc = objUserControl.Controls("DataFimRealCalc")
End Property

Public Property Get DataInicioRealCalc() As Object
     Set DataInicioRealCalc = objUserControl.Controls("DataInicioRealCalc")
End Property

Public Property Get PercCompRealCalc() As Object
     Set PercCompRealCalc = objUserControl.Controls("PercCompRealCalc")
End Property

Public Property Get UpDownDataInicioReal() As Object
     Set UpDownDataInicioReal = objUserControl.Controls("UpDownDataInicioReal")
End Property

Public Property Get DataInicioRealInf() As Object
     Set DataInicioRealInf = objUserControl.Controls("DataInicioRealInf")
End Property

Public Property Get UpDownDataFimReal() As Object
     Set UpDownDataFimReal = objUserControl.Controls("UpDownDataFimReal")
End Property

Public Property Get DataFimRealInf() As Object
     Set DataFimRealInf = objUserControl.Controls("DataFimRealInf")
End Property

Public Property Get PercCompRealInf() As Object
     Set PercCompRealInf = objUserControl.Controls("PercCompRealInf")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get Responsavel() As Object
     Set Responsavel = objUserControl.Controls("Responsavel")
End Property

Public Property Get Objetivo() As Object
     Set Objetivo = objUserControl.Controls("Objetivo")
End Property

Public Property Get Justificativa() As Object
     Set Justificativa = objUserControl.Controls("Justificativa")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get LabelCliente() As Object
     Set LabelCliente = objUserControl.Controls("LabelCliente")
End Property

Public Property Get optEmpToda() As Object
     Set optEmpToda = objUserControl.Controls("optEmpToda")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objUserControl.Controls("NomeReduzido")
End Property

Public Property Get UpDownDataCriacao() As Object
     Set UpDownDataCriacao = objUserControl.Controls("UpDownDataCriacao")
End Property

Public Property Get DataCriacao() As Object
     Set DataCriacao = objUserControl.Controls("DataCriacao")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get LabelCodigo() As Object
     Set LabelCodigo = objUserControl.Controls("LabelCodigo")
End Property

Public Property Get LabelNomeReduzido() As Object
     Set LabelNomeReduzido = objUserControl.Controls("LabelNomeReduzido")
End Property

Public Property Get StatusProjeto() As Object
     Set StatusProjeto = objUserControl.Controls("StatusProjeto")
End Property

Public Property Get BotaoDocRelacs() As Object
     Set BotaoDocRelacs = objUserControl.Controls("BotaoDocRelacs")
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

Public Property Get BotaoFluxo() As Object
     Set BotaoFluxo = objUserControl.Controls("BotaoFluxo")
End Property

Public Property Get BotaoContrato() As Object
     Set BotaoContrato = objUserControl.Controls("BotaoContrato")
End Property

Public Property Get BotaoCronograma() As Object
     Set BotaoCronograma = objUserControl.Controls("BotaoCronograma")
End Property

Public Property Get BotaoFisicoFin() As Object
     Set BotaoFisicoFin = objUserControl.Controls("BotaoFisicoFin")
End Property

Public Property Get BotaoProposta() As Object
     Set BotaoProposta = objUserControl.Controls("BotaoProposta")
End Property

Public Property Get BotaoOrganograma() As Object
     Set BotaoOrganograma = objUserControl.Controls("BotaoOrganograma")
End Property

Public Property Get Segmento() As Object
     Set Segmento = objUserControl.Controls("Segmento")
End Property
'--- fim dos properties get dos controles da tela

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Projetos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Projetos"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub UserControl_Initialize()
    Set gobjTelaEscopo = New ClassTelaEscopo
    Set gobjTelaCamposCust = New ClassTelaDadosCust
    Set gobjTelaEscopo.objUserControl = Me
    Set gobjTelaCamposCust.objUserControl = Me
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
        
    End If

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

Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set gobjTelaEscopo = Nothing
    Set gobjTelaCamposCust = Nothing
    
    Set gobjProjeto = Nothing
    Set gobjAnexos = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181589)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    Set gobjAnexos = New ClassAnexos

    Call gobjTelaCamposCust.Exibe_Campos_Customizados
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    lErro = Inicializa_Mascara_Projeto(Codigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Carrega_Usuarios(Responsavel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_PRJ_OBJETIVOS, Objetivo, True, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_PRJ_SEGMENTO, Segmento, True, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    iFrameAtual = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181590)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objProjetos As ClassProjetos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objProjetos Is Nothing) Then

        lErro = Traz_Projetos_Tela(objProjetos)
        If lErro <> SUCESSO Then gError 181319

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 181319

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181591)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189091

    objProjetos.sCodigo = sProjeto
    
    If optEmpToda.Value = vbChecked Then
        objProjetos.iFilialEmpresa = EMPRESA_TODA
    Else
        objProjetos.iFilialEmpresa = giFilialEmpresa
    End If
    objProjetos.sNomeReduzido = NomeReduzido.Text
    objProjetos.sDescricao = Descricao.Text
    objProjetos.dtDataCriacao = StrParaDate(DataCriacao.Text)
    
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 181433

        If lErro = SUCESSO Then
            objProjetos.lCliente = objCliente.lCodigo
        End If
            
    End If

    objProjetos.iFilialCliente = Codigo_Extrai(Filial.Text)
    
    objProjetos.sResponsavel = Responsavel.Text
    objProjetos.sObjetivo = Objetivo.Text
    objProjetos.sSegmento = Segmento.Text
    objProjetos.sJustificativa = Justificativa.Text
    objProjetos.sObservacao = Observacao.Text
    
    objProjetos.dtDataInicio = StrParaDate(DataInicioInf.Text)
    objProjetos.dtDataFim = StrParaDate(DataFimInf.Text)
    objProjetos.dtDataInicioReal = StrParaDate(DataInicioRealInf.Text)
    objProjetos.dtDataFimReal = StrParaDate(DataFimRealInf.Text)
    
    objProjetos.dPercentualComplet = StrParaDbl(Val(PercCompRealInf.Text) / 100)
    
    lErro = gobjTelaEscopo.Move_Tela_Memoria(Me, objProjetos.objEscopo)
    If lErro <> SUCESSO Then gError 181571
    
    lErro = gobjTelaCamposCust.Move_Tela_Memoria(objProjetos.objCamposCust, objProjetos.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 181572
    
    Set objProjetos.objAnexos = gobjAnexos

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 181433, 181571, 181572, 189091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181592)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objProjetos As New ClassProjetos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Projetos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objProjetos)
    If lErro <> SUCESSO Then gError 181320

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProjetos.sCodigo, STRING_PRJ_CODIGO, "Codigo"
    'Filtros para o Sistema de Setas
    
    colSelecao.Add "(FilialEmpresa = 0 OR FilialEmpresa = " & CStr(giFilialEmpresa) & ") AND 1", OP_IGUAL, 1
    
    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 181320

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181593)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objProjetos As New ClassProjetos

On Error GoTo Erro_Tela_Preenche

    objProjetos.sCodigo = colCampoValor.Item("Codigo").vValor
    objProjetos.iFilialEmpresa = giFilialEmpresa

    If objProjetos.sCodigo <> "" Then
        lErro = Traz_Projetos_Tela(objProjetos)
        If lErro <> SUCESSO Then gError 181321
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 181321

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181594)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProjetos As New ClassProjetos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 181583
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 181584
    If Len(Trim(DataCriacao.ClipText)) = 0 Then gError 181585

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objProjetos)
    If lErro <> SUCESSO Then gError 181322
    
    If objProjetos.dtDataInicio <> DATA_NULA And objProjetos.dtDataFim <> DATA_NULA Then
        If objProjetos.dtDataInicio > objProjetos.dtDataFim Then gError 181586
    End If
    
    If objProjetos.dtDataInicioReal <> DATA_NULA And objProjetos.dtDataFimReal <> DATA_NULA Then
        If objProjetos.dtDataInicioReal > objProjetos.dtDataFimReal Then gError 181587
    End If

    lErro = Trata_Alteracao(objProjetos, objProjetos.sCodigo, objProjetos.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 181323
    
    Set objProjetos.objTela = Me

    'Grava o/a Projetos no Banco de Dados
    lErro = CF("Projetos_Grava", objProjetos)
    If lErro <> SUCESSO Then gError 181324

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 181322, 181323, 181324
        
        Case 181583
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            'Codigo.SetFocus

        Case 181584
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_PRJ_NAO_PREENCHIDO", gErr)
            'NomeReduzido.SetFocus
            
        Case 181585
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACRIACAO_PRJ_NAO_PREENCHIDO", gErr)
            'DataCriacao.SetFocus
            
        Case 181586
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            'DataInicioInf.SetFocus
            
        Case 181587
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            'DataInicioRealInf.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181595)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Projetos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Projetos

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Set gobjAnexos = New ClassAnexos

    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    DataInicioCalc.Caption = ""
    DataFimCalc.Caption = ""
    DataInicioRealCalc.Caption = ""
    DataFimRealCalc.Caption = ""
    PercCompRealCalc.Caption = ""
    
    Responsavel.Text = ""
    Objetivo.Text = ""
    Segmento.Text = ""
    RespUsu.Value = vbUnchecked
    
    optEmpToda.Value = vbUnchecked
    
    Filial.ListIndex = -1
    
    Call CF2(Me, "PRJ_Limpa_Tela_Projetos")

    iAlterado = 0

    Limpa_Tela_Projetos = SUCESSO

    Exit Function

Erro_Limpa_Tela_Projetos:

    Limpa_Tela_Projetos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181596)

    End Select

    Exit Function

End Function

Function Traz_Projetos_Tela(objProjetos As ClassProjetos) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Projetos_Tela

    Call Limpa_Tela_Projetos

    'Lê o Projetos que está sendo Passado
    lErro = CF("Projetos_Le", objProjetos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181325

    lErro = Retorno_Projeto_Tela(Codigo, objProjetos.sCodigo)
    If lErro <> SUCESSO Then gError 189130
    
    NomeReduzido.Text = objProjetos.sNomeReduzido

    If lErro = SUCESSO Then
    
        If objProjetos.iFilialEmpresa = EMPRESA_TODA Then
            optEmpToda.Value = vbChecked
        Else
            optEmpToda.Value = vbUnchecked
        End If
    
        lErro = CF("PRJEtapas_Le_Projeto", objProjetos)
        If lErro <> SUCESSO Then gError 189255

        Call objProjetos.Calcula_Dados_Calculados
        
        Descricao.Text = objProjetos.sDescricao

        If objProjetos.dtDataCriacao <> DATA_NULA Then
            DataCriacao.PromptInclude = False
            DataCriacao.Text = Format(objProjetos.dtDataCriacao, "dd/mm/yy")
            DataCriacao.PromptInclude = True
        End If

        If objProjetos.lCliente <> 0 Then

            Cliente.Text = CStr(objProjetos.lCliente)
            Call Cliente_Validate(bSGECancelDummy)
            
            Filial.Text = objProjetos.iFilialCliente
            Call Filial_Validate(bSGECancelDummy)
        
        End If

        Responsavel.Text = objProjetos.sResponsavel
        Call Responsavel_Validate(bSGECancelDummy)
        
        Objetivo.Text = objProjetos.sObjetivo
        Segmento.Text = objProjetos.sSegmento
        Justificativa.Text = objProjetos.sJustificativa
        Observacao.Text = objProjetos.sObservacao

        If objProjetos.dtDataInicio <> DATA_NULA Then
            DataInicioInf.PromptInclude = False
            DataInicioInf.Text = Format(objProjetos.dtDataInicio, "dd/mm/yy")
            DataInicioInf.PromptInclude = True
        End If

        If objProjetos.dtDataFim <> DATA_NULA Then
            DataFimInf.PromptInclude = False
            DataFimInf.Text = Format(objProjetos.dtDataFim, "dd/mm/yy")
            DataFimInf.PromptInclude = True
        End If
        
        Intervalo.PromptInclude = False
        Intervalo.Text = DateDiff("d", objProjetos.dtDataInicio, objProjetos.dtDataFim) + 1
        Intervalo.PromptInclude = True

        If objProjetos.dtDataInicioReal <> DATA_NULA Then
            DataInicioRealInf.PromptInclude = False
            DataInicioRealInf.Text = Format(objProjetos.dtDataInicioReal, "dd/mm/yy")
            DataInicioRealInf.PromptInclude = True
        End If

        If objProjetos.dtDataFimReal <> DATA_NULA Then
            DataFimRealInf.PromptInclude = False
            DataFimRealInf.Text = Format(objProjetos.dtDataFimReal, "dd/mm/yy")
            DataFimRealInf.PromptInclude = True
        End If

        PercCompRealInf.Text = CStr(objProjetos.dPercentualComplet * 100)
        
        If objProjetos.objDadosCalc.dtDataIniPrev <> DATA_NULA Then
            DataInicioCalc.Caption = Format(objProjetos.objDadosCalc.dtDataIniPrev, "dd/mm/yyyy")
        End If

        If objProjetos.objDadosCalc.dtDataFimPrev <> DATA_NULA Then
            DataFimCalc.Caption = Format(objProjetos.objDadosCalc.dtDataFimPrev, "dd/mm/yyyy")
        End If

        If objProjetos.objDadosCalc.dtDataIniReal <> DATA_NULA Then
            DataInicioRealCalc.Caption = Format(objProjetos.objDadosCalc.dtDataIniReal, "dd/mm/yyyy")
        End If

        If objProjetos.objDadosCalc.dtDataFimReal <> DATA_NULA Then
            DataFimRealCalc.Caption = Format(objProjetos.objDadosCalc.dtDataFimReal, "dd/mm/yyyy")
        End If

        PercCompRealCalc.Caption = Format(objProjetos.objDadosCalc.dPercentualComplet, "PERCENT")
        
        objProjetos.objEscopo.lNumIntDoc = objProjetos.lNumIntDocEscopo

        lErro = gobjTelaEscopo.Traz_PRJEscopo_Tela(Me, objProjetos.objEscopo)
        If lErro <> SUCESSO Then gError 181574
        
        objProjetos.objCamposCust.iTipoNumIntDocOrigem = CAMPO_CUSTOMIZADO_TIPO_PROJETO
        objProjetos.objCamposCust.lNumIntDocOrigem = objProjetos.lNumIntDoc

        lErro = gobjTelaCamposCust.Traz_CamposCustomizados_Tela(objProjetos.objCamposCust)
        If lErro <> SUCESSO Then gError 181575
        
        lErro = CF2(Me, "PRJ_Traz_Projetos_Tela", objProjetos)
        If lErro <> SUCESSO Then gError 181575
        
        Set gobjAnexos = objProjetos.objAnexos
        If gobjAnexos Is Nothing Then
            Set gobjAnexos = New ClassAnexos
            gobjAnexos.iTipoDoc = ANEXO_TIPO_PRJ
            gobjAnexos.lNumIntDoc = objProjetos.lNumIntDoc
        End If
        lErro = CF("Anexos_Le", gobjAnexos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Traz_Projetos_Tela = SUCESSO

    Exit Function

Erro_Traz_Projetos_Tela:

    Traz_Projetos_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 181325, 181574, 181575, 189130, 189255

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181597)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 181326

    'Limpa Tela
    Call Limpa_Tela_Projetos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 181326

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181598)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181599)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 181327

    Call Limpa_Tela_Projetos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 181327

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181600)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objProjetos As New ClassProjetos
Dim vbMsgRes As VbMsgBoxResult
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 181588
    
    lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189092
    
    objProjetos.sCodigo = sProjeto
    objProjetos.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PROJETOS", objProjetos.sCodigo, objProjetos.iFilialEmpresa)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Projetos_Exclui", objProjetos)
        If lErro <> SUCESSO Then gError 181328

        'Limpa Tela
        Call Limpa_Tela_Projetos

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 181328, 189092
        
        Case 181588
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181601)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.ClipText)) <> 0 Then

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 181329

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181602)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Public Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Verifica se NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

       '#######################################
       'CRITICA NomeReduzido
       '#######################################

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181603)

    End Select

    Exit Sub

End Sub

Public Sub NomeReduzido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181604)

    End Select

    Exit Sub

End Sub

Public Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownDataCriacao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_DownClick

    DataCriacao.SetFocus

    If Len(DataCriacao.ClipText) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181330

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_DownClick:

    Select Case gErr

        Case 181330

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181605)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCriacao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_UpClick

    DataCriacao.SetFocus

    If Len(Trim(DataCriacao.ClipText)) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181331

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_UpClick:

    Select Case gErr

        Case 181331

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181606)

    End Select

    Exit Sub

End Sub

Public Sub DataCriacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCriacao, iAlterado)
    
End Sub

Public Sub DataCriacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCriacao_Validate

    If Len(Trim(DataCriacao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCriacao.Text)
        If lErro <> SUCESSO Then gError 181332

    End If

    Exit Sub

Erro_DataCriacao_Validate:

    Cancel = True

    Select Case gErr

        Case 181332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181607)

    End Select

    Exit Sub

End Sub

Public Sub DataCriacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome


On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 181635
                   
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 181636

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 181635, 181636

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181608)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Public Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 181637

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 181638

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 181639

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 181640

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 181641
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        Else
            
            objCliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 181642
        
        If lErro = SUCESSO Then gError 181643
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 181637, 181639

        Case 181638
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 181640, 181642 'tratado na rotina chamada

        Case 181641
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 181643
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181609)

    End Select

    Exit Sub


End Sub

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Responsavel_Validate

    RespUsu.Value = vbUnchecked

    'Verifica se Responsavel está preenchida
    If Len(Trim(Responsavel.Text)) <> 0 Then

        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = Responsavel.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then RespUsu.Value = vbChecked

    End If

    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185161)

    End Select

    Exit Sub

End Sub

Public Sub Responsavel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Objetivo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Objetivo_Validate

    'Verifica se Objetivo está preenchida
    If Len(Trim(Objetivo.Text)) <> 0 Then

       '#######################################
       'CRITICA Objetivo
       '#######################################

    End If

    Exit Sub

Erro_Objetivo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181611)

    End Select

    Exit Sub

End Sub

Public Sub Objetivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Segmento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Segmento_Validate

    'Verifica se Segmento está preenchida
    If Len(Trim(Segmento.Text)) <> 0 Then

       '#######################################
       'CRITICA Segmento
       '#######################################

    End If

    Exit Sub

Erro_Segmento_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181611)

    End Select

    Exit Sub

End Sub

Public Sub Segmento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Justificativa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Justificativa_Validate

    'Verifica se Justificativa está preenchida
    If Len(Trim(Justificativa.Text)) <> 0 Then

       '#######################################
       'CRITICA Justificativa
       '#######################################

    End If

    Exit Sub

Erro_Justificativa_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181612)

    End Select

    Exit Sub

End Sub

Public Sub Justificativa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181613)

    End Select

    Exit Sub

End Sub

Public Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    DataInicioInf.SetFocus

    If Len(DataInicioInf.ClipText) > 0 Then

        sData = DataInicioInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181335

        DataInicioInf.Text = sData
        
        Call DataInicioInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 181335

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181614)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    DataInicioInf.SetFocus

    If Len(Trim(DataInicioInf.ClipText)) > 0 Then

        sData = DataInicioInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181336

        DataInicioInf.Text = sData
        
        Call DataInicioInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 181336

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181615)

    End Select

    Exit Sub

End Sub

Public Sub DataInicioInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioInf, iAlterado)
    
End Sub

Public Sub DataInicioInf_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData1 As Date
Dim dtData2 As Date

On Error GoTo Erro_DataInicioInf_Validate

    If Len(Trim(DataInicioInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicioInf.Text)
        If lErro <> SUCESSO Then gError 181337

        dtData1 = StrParaDate(DataInicioInf.Text)

        If StrParaDate(DataFimInf.Text) <> DATA_NULA Then
        
            dtData2 = StrParaDate(DataFimInf.Text)
            
            Intervalo.PromptInclude = False
            Intervalo.Text = CStr(1 + DateDiff("d", dtData1, dtData2))
            Intervalo.PromptInclude = True
        
        Else
        
            If Len(Trim(Intervalo.Text)) <> 0 Then
        
                DataFimInf.PromptInclude = False
                DataFimInf.Text = Format(DateAdd("d", StrParaInt(Intervalo.Text) - 1, dtData1), "dd/mm/yy")
                DataFimInf.PromptInclude = True
        
            End If
            
        End If
    
    End If

    Exit Sub

Erro_DataInicioInf_Validate:

    Cancel = True

    Select Case gErr

        Case 181337

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181616)

    End Select

    Exit Sub

End Sub

Public Sub DataInicioInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_DownClick

    DataFimInf.SetFocus

    If Len(DataFimInf.ClipText) > 0 Then

        sData = DataFimInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181338

        DataFimInf.Text = sData
        
        Call DataFimInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 181338

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181617)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_UpClick

    DataFimInf.SetFocus

    If Len(Trim(DataFimInf.ClipText)) > 0 Then

        sData = DataFimInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181339

        DataFimInf.Text = sData
        
        Call DataFimInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 181339

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181618)

    End Select

    Exit Sub

End Sub

Public Sub DataFimInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimInf, iAlterado)
    
End Sub

Public Sub DataFimInf_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData1 As Date
Dim dtData2 As Date

On Error GoTo Erro_DataFimInf_Validate

    If Len(Trim(DataFimInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFimInf.Text)
        If lErro <> SUCESSO Then gError 181340

        dtData2 = StrParaDate(DataFimInf.Text)

        If StrParaDate(DataInicioInf.Text) <> DATA_NULA Then
        
            dtData1 = StrParaDate(DataInicioInf.Text)
            
            Intervalo.PromptInclude = False
            Intervalo.Text = CStr(1 + DateDiff("d", dtData1, dtData2))
            Intervalo.PromptInclude = True
        
        Else
        
            If Len(Trim(Intervalo.Text)) <> 0 Then
        
                DataInicioInf.PromptInclude = False
                DataInicioInf.Text = Format(DateAdd("d", -(StrParaInt(Intervalo.Text) - 1), dtData2), "dd/mm/yy")
                DataInicioInf.PromptInclude = True
        
            End If
            
        End If
        
    End If

    Exit Sub

Erro_DataFimInf_Validate:

    Cancel = True

    Select Case gErr

        Case 181340

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181619)

    End Select

    Exit Sub

End Sub

Public Sub DataFimInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownDataInicioReal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicioReal_DownClick

    DataInicioRealInf.SetFocus

    If Len(DataInicioRealInf.ClipText) > 0 Then

        sData = DataInicioRealInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181341

        DataInicioRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicioReal_DownClick:

    Select Case gErr

        Case 181341

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181620)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataInicioReal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicioReal_UpClick

    DataInicioRealInf.SetFocus

    If Len(Trim(DataInicioRealInf.ClipText)) > 0 Then

        sData = DataInicioRealInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181342

        DataInicioRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicioReal_UpClick:

    Select Case gErr

        Case 181342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181621)

    End Select

    Exit Sub

End Sub

Public Sub DataInicioRealInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioRealInf, iAlterado)
    
End Sub

Public Sub DataInicioRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicioRealInf_Validate

    If Len(Trim(DataInicioRealInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicioRealInf.Text)
        If lErro <> SUCESSO Then gError 181343

    End If

    Exit Sub

Erro_DataInicioRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 181343

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181622)

    End Select

    Exit Sub

End Sub

Public Sub DataInicioRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownDataFimReal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFimReal_DownClick

    DataFimRealInf.SetFocus

    If Len(DataFimRealInf.ClipText) > 0 Then

        sData = DataFimRealInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181344

        DataFimRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFimReal_DownClick:

    Select Case gErr

        Case 181344

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181623)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataFimReal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFimReal_UpClick

    DataFimRealInf.SetFocus

    If Len(Trim(DataFimRealInf.ClipText)) > 0 Then

        sData = DataFimRealInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181345

        DataFimRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFimReal_UpClick:

    Select Case gErr

        Case 181345

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181624)

    End Select

    Exit Sub

End Sub

Public Sub DataFimRealInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimRealInf, iAlterado)
    
End Sub

Public Sub DataFimRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFimRealInf_Validate

    If Len(Trim(DataFimRealInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFimRealInf.Text)
        If lErro <> SUCESSO Then gError 181346

        PercCompRealInf.Text = "100"

    End If

    Exit Sub

Erro_DataFimRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 181346

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181625)

    End Select

    Exit Sub

End Sub

Public Sub DataFimRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercCompRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercCompRealInf_Validate

    'Verifica se PercCompRealInf está preenchida
    If Len(Trim(PercCompRealInf.Text)) <> 0 Then

       'Critica a PercCompRealInf
       lErro = Porcentagem_Critica(PercCompRealInf.Text)
       If lErro <> SUCESSO Then gError 181347

    End If

    Exit Sub

Erro_PercCompRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 181347

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181626)

    End Select

    Exit Sub

End Sub

Public Sub PercCompRealInf_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercCompRealInf, iAlterado)
End Sub

Public Sub PercCompRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LabelCodigo_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189093

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoCodigo, , "Código")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case 189093

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Sub

End Sub

Public Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

        objProjeto.sNomeReduzido = NomeReduzido.Text

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoCodigo, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181628)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objProjeto = obj1

    'Mostra os dados do CentrodeTrabalho na tela
    lErro = Traz_Projetos_Tela(objProjeto)
    If lErro <> SUCESSO Then gError 134369
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 134369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181629)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


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

'##################################################################
'Tem que colocar o código para o modo de edição aqui







'##################################################################

'##################################################################
'Tratamento dos Campos customizados
Public Sub Data_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Data_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Data_GotFocus(Index)
End Sub

Public Sub Data_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Data_Validate(Index, Cancel)
End Sub

Public Sub Numero_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Numero_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Numero_GotFocus(Index)
End Sub

Public Sub Numero_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Numero_Validate(Index, Cancel)
End Sub

Public Sub Valor_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Valor_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Valor_GotFocus(Index)
End Sub

Public Sub Valor_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Valor_Validate(Index, Cancel)
End Sub

Public Sub Texto_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UpDownData_DownClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_DownClick(Index)
End Sub

Public Sub UpDownData_UpClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_UpClick(Index)
End Sub

Public Sub BotaoDadosCustNovo_Click()
    Call gobjTelaCamposCust.BotaoDadosCustNovo_Click
End Sub

Public Sub BotaoDadosCustDel_Click()
    Call gobjTelaCamposCust.BotaoDadosCustDel_Click
End Sub
'##################################################################

'##################################################################
'Tratamento do Escopo
Public Sub EscDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EscExclusoes_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EscExpectativa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EscFatores_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EscPremissas_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EscRestricoes_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'##################################################################

'##################################################################
'Tratamento Tab de Datas
Public Sub Intervalo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Intervalo_Validate(Cancel As Boolean)

Dim dtData As Date

    If Len(Trim(Intervalo)) <> 0 Then

        If StrParaDate(DataInicioInf.Text) <> DATA_NULA Then
                
            dtData = StrParaDate(DataInicioInf.Text)
            
            DataFimInf.PromptInclude = False
            DataFimInf.Text = Format(DateAdd("d", StrParaInt(Intervalo.Text) - 1, dtData), "dd/mm/yy")
            DataFimInf.PromptInclude = True
        
        Else
        
            If StrParaDate(DataFimInf.Text) <> DATA_NULA Then
    
                dtData = StrParaDate(DataFimInf.Text)
    
                DataInicioInf.PromptInclude = False
                DataInicioInf.Text = Format(DateAdd("d", -(StrParaInt(Intervalo.Text) - 1), dtData), "dd/mm/yy")
                DataInicioInf.PromptInclude = True
        
            End If
            
        End If
        
    End If
    
End Sub
'##################################################################

Public Sub BotaoOrganograma_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_BotaoOrganograma_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 185015
    
    lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189094
    
    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa

    'Lê o Projetos que está sendo Passado
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185016
    
    If lErro <> SUCESSO Then gError 185017
    
    Call Chama_Tela("OrganogramaPRJ", objProjeto)

    Exit Sub

Erro_BotaoOrganograma_Click:

    Select Case gErr

        Case 185015
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 185016, 189094
        
        Case 185017
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185018)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoDocRelacs_Click()

Dim lErro As Long
Dim objProjetoInfo As New ClassProjetoInfo
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoDocRelacs_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 185019
    
    lErro = Projeto_Formata(Codigo.Text, sProjeto, iPreenchido)
    If lErro <> SUCESSO Then gError 189295
    
    colSelecao.Add sProjeto

    Call Chama_Tela("ProjetoInfoLista", colSelecao, objProjetoInfo, Nothing, "Projeto = ?")

    Exit Sub

Erro_BotaoDocRelacs_Click:

    Select Case gErr

        Case 185019
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 189295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181627)

    End Select

    Exit Sub

End Sub

Public Sub BotaoCronograma_Click()
'Abre a tela de Cronograma Gráfico

Dim lErro As Long
Dim objTelaGrafico As New ClassTelaGrafico

On Error GoTo Erro_BotaoCronograma_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico)
    If lErro <> SUCESSO Then gError 185541
    
    Call Chama_Tela("TelaGrafico", objTelaGrafico)

    Exit Sub

Erro_BotaoCronograma_Click:

    Select Case gErr
    
        Case 185541

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185542)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoFisicoFin_Click()
'Abre a tela de FisicoFin Gráfico

Dim lErro As Long
Dim objTelaGrafico As New ClassTelaGrafico

On Error GoTo Erro_BotaoFisicoFin_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico, True)
    If lErro <> SUCESSO Then gError 185541
    
    Call Chama_Tela("TelaGrafico", objTelaGrafico)

    Exit Sub

Erro_BotaoFisicoFin_Click:

    Select Case gErr
    
        Case 185541

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185542)

    End Select
    
    Exit Sub
    
End Sub

Public Function Atualiza_Cronograma(objTelaGrafico As ClassTelaGrafico, Optional ByVal bFisicoFinac As Boolean = False, Optional ByVal sNomeArqSalvar As String = "") As Long
'Acerta a tela de cronograma gráfico após o retorno da tela de etapas
'Para isso remonta o objTelaGrafico com os novos dados

Dim lErro As Long
Dim objTelaGraficoItem As New ClassTelaGraficoItens
Dim objTelaGraficoItemPre As New ClassTelaGraficoItens
Dim objProjeto As New ClassProjetos
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim objEtapaAux2 As ClassPRJEtapas
Dim objEtapaPre As ClassPRJEtapasPredecessoras
Dim objBotao As ClassTelaGraficoBotao
Dim objPRJRecebPagto As New ClassPRJRecebPagto
Dim objTelaGraficoParam As ClassTelaGraficoParam
Dim objTela As Object
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Atualiza_Cronograma:

    If objUserControl Is Nothing Then
    
        objProjeto.lNumIntDoc = objTelaGrafico.colItens.Item(1).colobj.Item(1).lNumIntDocPRJ
    
    Else

        If Len(Trim(Codigo.ClipText)) = 0 Then gError 185550
        
        lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189096
        
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
    End If

    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185551
        
    objTelaGrafico.dtDataInicio = DATA_NULA
    
    lErro = CF("PRJEtapas_Le_Projeto", objProjeto)
    If lErro <> SUCESSO Then gError 185552
    
    Set gobjProjeto = objProjeto

    Set objTelaGrafico.colBotoes = New Collection
    Set objTelaGrafico.colItens = New Collection
    Set objTelaGrafico.colParametros = New Collection

    Set objTelaGrafico.objTela = Me
    
    objTelaGrafico.sNomeArqFigura = sNomeArqSalvar
    
    objPRJRecebPagto.lNumIntDocPRJ = objProjeto.lNumIntDoc
    
    objTelaGrafico.sNomeTela = "Cronograma das Etapas do Projeto"
    objTelaGrafico.iTamanhoDia = 540
    objTelaGrafico.iModal = DESMARCADO
    objTelaGrafico.iAtualizaRetornoClick = MARCADO
    objTelaGrafico.sNomeFuncAtualiza = "Atualiza_Cronograma"
    objTelaGrafico.iAlteraQtdLinhasPara = 12
    objTelaGrafico.iCadaEtapaUmaLinha = MARCADO
    objTelaGrafico.sTextoImpressao = "Projeto: " & objProjeto.sCodigo & vbNewLine & "Descrição: " & objProjeto.sDescricao & vbNewLine & "Responsável: " & objProjeto.sResponsavel

    objTelaGrafico.colParametros.Add bFisicoFinac
    
    If bFisicoFinac Then
    
        objTelaGrafico.sNomeTela = "Cronograma Físico Financeiro do Projeto"
        objTelaGrafico.iAlteraQtdLinhasPara = 7
        objTelaGrafico.iExibirGridAux = MARCADO
        objTelaGrafico.sNomeFuncTrataGrid = "Cronograma_Trata_GridAux_PRJ"
        objTelaGrafico.iAlturaDia = 240
        objTelaGrafico.iTamanhoDia = 810
        
        Set objTela = Me
        
        Set objTelaGraficoParam = New ClassTelaGraficoParam
        
        objTelaGraficoParam.sNomeFuncao = TELAGRAFICO_FUNCAO_PREENCHE_GRIDAUX
        Set objTelaGraficoParam.vParam = objProjeto
        
        objTelaGrafico.colParametrosTrataGrid.Add objTelaGraficoParam
    
        Set objTelaGraficoParam = New ClassTelaGraficoParam
        
        objTelaGraficoParam.sNomeFuncao = TELAGRAFICO_FUNCAO_PREENCHE_GRIDAUX
        Set objTelaGraficoParam.vParam = objTela
        
        objTelaGrafico.colParametrosTrataGrid.Add objTelaGraficoParam
    
        Set objBotao = New ClassTelaGraficoBotao
    
        objBotao.colParametros.Add objPRJRecebPagto
        objBotao.sNome = "Pagto"
        objBotao.sNomeFuncao = "Cronograma_Adiciona_Pagto"
        objBotao.sTextoExibicao = "Pagamento"
        objBotao.iAtualizaRetornoClick = MARCADO
        
        objTelaGrafico.colBotoes.Add objBotao
        
        Set objBotao = New ClassTelaGraficoBotao
    
        objBotao.colParametros.Add objPRJRecebPagto
        objBotao.sNome = "Receb"
        objBotao.sNomeFuncao = "Cronograma_Adiciona_Receb"
        objBotao.sTextoExibicao = "Recebimento"
        objBotao.iAtualizaRetornoClick = MARCADO
        
        objTelaGrafico.colBotoes.Add objBotao
        
'        Set objBotao = New ClassTelaGraficoBotao
'
'        objBotao.sNome = "Modificar Pagtos"
'        objBotao.sNomeFuncao = "Cronograma_Modifica_Pagto"
'        objBotao.sTextoExibicao = "Modificar Pagtos"
'        objBotao.iAtualizaRetornoClick = MARCADO
'
'        objTelaGrafico.colBotoes.Add objBotao
'
'        Set objBotao = New ClassTelaGraficoBotao
'
'        objBotao.sNome = "Modificar Receb."
'        objBotao.sNomeFuncao = "Cronograma_Modifica_Receb"
'        objBotao.sTextoExibicao = "Modificar Receb."
'        objBotao.iAtualizaRetornoClick = MARCADO
'
'        objTelaGrafico.colBotoes.Add objBotao
        
        Set objBotao = New ClassTelaGraficoBotao
    
        objBotao.sNome = "Gráfico Excel"
        objBotao.sNomeFuncao = "Cronograma_Grafico_Excel"
        objBotao.sTextoExibicao = "Gráfico Excel"
        objBotao.iAtualizaRetornoClick = MARCADO
        
        objTelaGrafico.colBotoes.Add objBotao
        
    End If

    'Para cada etapa do projeto
    For Each objEtapa In objProjeto.colEtapas
    
        If objEtapa.dtDataFim <> DATA_NULA And objEtapa.dtDataInicio <> DATA_NULA Then
        
            Set objTelaGraficoItem = New ClassTelaGraficoItens
        
            objTelaGraficoItem.colobj.Add objEtapa
                        
            objTelaGraficoItem.dtDataFim = objEtapa.dtDataFim
            objTelaGraficoItem.dtDataInicio = objEtapa.dtDataInicio
            objTelaGraficoItem.sNomeTela = "EtapaPRJ"
            objTelaGraficoItem.sTextoExibicao = "Projeto: " & objProjeto.sNomeReduzido & vbNewLine & "Etapa: " & objEtapa.sCodigo & vbNewLine & "Nome: " & objEtapa.sNomeReduzido & vbNewLine & "Data Início Prev: " & Format(objEtapa.dtDataInicio, "dd/mm/yyyy") & vbNewLine & "Data Fim Prev: " & Format(objEtapa.dtDataFim, "dd/mm/yyyy") & vbNewLine & "% Completado: " & Format(objEtapa.dPercentualComplet, "PERCENT")
            objTelaGraficoItem.sDescricao = objEtapa.sNomeReduzido
            objTelaGraficoItem.sNome = objEtapa.sCodigo
            objTelaGraficoItem.iIndiceCor = 1
                
            objTelaGrafico.colItens.Add objTelaGraficoItem
            
        End If
        
    Next
    
    For Each objTelaGraficoItem In objTelaGrafico.colItens
    
        Set objEtapa = objTelaGraficoItem.colobj.Item(1)
        
        Set objEtapaAux2 = New ClassPRJEtapas
        
        objEtapaAux2.lNumIntDoc = objEtapa.lNumIntDoc
        objEtapaAux2.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ
        objEtapaAux2.sCodigo = objEtapa.sCodigo
        
        lErro = CF("PRJEtapasPredecessoras_Le", objEtapaAux2)
        If lErro <> SUCESSO Then gError 185740
        
        For Each objEtapaPre In objEtapaAux2.colPredecessores
        
            For Each objTelaGraficoItemPre In objTelaGrafico.colItens
            
                Set objEtapaAux = objTelaGraficoItemPre.colobj.Item(1)
                
                If objEtapaAux.lNumIntDoc = objEtapaPre.lNumIntDocEtapaPre Then
                
                    objTelaGraficoItem.colPredecessores.Add objTelaGraficoItemPre
                    
                    Exit For
                
                End If
            
            Next
            
        Next
    
    Next
    
    Atualiza_Cronograma = SUCESSO

    Exit Function

Erro_Atualiza_Cronograma:

    Atualiza_Cronograma = gErr

    Select Case gErr
    
        Case 185550
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 185551, 185552, 189096
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185553)

    End Select
    
    Exit Function
    
End Function

Public Function Cronograma_Adiciona_Pagto(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJRecebPagtoAux As New ClassPRJRecebPagto

On Error GoTo Erro_Cronograma_Adiciona_Pagto

    objPRJRecebPagtoAux.iTipo = PRJ_TIPO_PAGTO
    objPRJRecebPagtoAux.lNumIntDocPRJ = objPRJRecebPagto.lNumIntDocPRJ

    lErro = CF("PRJRecebPagto_Le_CFF", objPRJRecebPagtoAux)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187492

    Call Chama_Tela("PagamentoPRJ", objPRJRecebPagtoAux)

    Cronograma_Adiciona_Pagto = SUCESSO

    Exit Function

Erro_Cronograma_Adiciona_Pagto:

    Cronograma_Adiciona_Pagto = gErr

    Select Case gErr
    
        Case 187492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185600)

    End Select

    Exit Function

End Function

Public Function Cronograma_Adiciona_Receb(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJRecebPagtoAux As New ClassPRJRecebPagto

On Error GoTo Erro_Cronograma_Adiciona_Receb

    objPRJRecebPagtoAux.iTipo = PRJ_TIPO_RECEB
    objPRJRecebPagtoAux.lNumIntDocPRJ = objPRJRecebPagto.lNumIntDocPRJ

    lErro = CF("PRJRecebPagto_Le_CFF", objPRJRecebPagtoAux)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187493
    
    Call Chama_Tela("RecebimentoPRJ", objPRJRecebPagtoAux)

    Cronograma_Adiciona_Receb = SUCESSO

    Exit Function

Erro_Cronograma_Adiciona_Receb:

    Cronograma_Adiciona_Receb = gErr

    Select Case gErr
    
        Case 187493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185601)

    End Select

    Exit Function

End Function

'Public Function Cronograma_Modifica_Pagto() As Long
'
'Dim lErro As Long
'Dim colSelecao As New Collection
'Dim objPRJPagto As New ClassPRJRecebPagto
'
'On Error GoTo Erro_Cronograma_Modifica_Pagto
'
'    Call Chama_Tela("PagamentoPRJLista", colSelecao, objPRJPagto, Nothing)
'
'    Cronograma_Modifica_Pagto = SUCESSO
'
'    Exit Function
'
'Erro_Cronograma_Modifica_Pagto:
'
'    Cronograma_Modifica_Pagto = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185602)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Cronograma_Modifica_Receb() As Long
'
'Dim lErro As Long
'Dim colSelecao As New Collection
'Dim objPRJReceb As New ClassPRJRecebPagto
'
'On Error GoTo Erro_Cronograma_Modifica_Receb
'
'    Call Chama_Tela("RecebimentoPRJLista", colSelecao, objPRJReceb, Nothing)
'
'    Cronograma_Modifica_Receb = SUCESSO
'
'    Exit Function
'
'Erro_Cronograma_Modifica_Receb:
'
'    Cronograma_Modifica_Receb = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185602)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Cronograma_Grafico_Excel() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim dLarguraColuna As Double
Dim sNome As String
Dim objPlanilha As New ClassPlanilhaExcel
Dim objRecebPagto As ClassPRJRecebPagto
Dim objRecebPagtoRegra As ClassPRJRecebPagtoRegras
Dim colCampos As New Collection
Dim colEntrada As New Collection
Dim colRecebPagtoRegra As New Collection
Dim dSaldoAcumulado As Double
Dim objParc As ClassCondicaoPagtoParc
Dim iParcela As Integer
Dim objRecebPagtoRegraAux As ClassPRJRecebPagtoRegras

On Error GoTo Erro_Cronograma_Grafico_Excel
    
    'Guarda o nome que será exibido para a planilha
    objPlanilha.sNomeGrafico = "Gráfico"
    objPlanilha.sNomePlanilha = "Cronograma"
    objPlanilha.iModoGravacao = DESMARCADO
    
    'Obtém a largura da coluna no padrão do Excel
    lErro = CF("Excel_Obtem_Largura_Coluna", 1400, dLarguraColuna)
    If lErro <> SUCESSO Then gError 187457
       
    For iIndice = 1 To 8
    
        'Instancia um novo obj para armazenar os dados da coluna
        Set objColunas = New ClassColunasExcel
        
        If iIndice = 3 Or iIndice = 8 Then
            'Guarda no objeto a largura dessa coluna
            objColunas.dLarguraColuna = dLarguraColuna * 3
        ElseIf iIndice = 1 Then
            'Guarda no objeto a largura dessa coluna
            objColunas.dLarguraColuna = dLarguraColuna * 2
        Else
            'Guarda no objeto a largura dessa coluna
            objColunas.dLarguraColuna = dLarguraColuna
        End If
        
        'Adiciona a coluna à coleção de colunas
        objPlanilha.colColunas.Add objColunas
    
        'Instancia um novo objeto para armazenar os dados da primeira célula da coluna em questão
        Set objCelulas = New ClassCelulasExcel
        
        Select Case iIndice
        
            Case 1
                sNome = "Documento"
            Case 2
                sNome = "Parcela"
            Case 3
                sNome = "Regra"
            Case 4
                sNome = "Data"
                objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
'                objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_LINE_MARKERS
'                objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS 'EXCEL_EXIBE_LABELS_VALOR
            Case 5
                sNome = "Receita"
                objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
                objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS 'EXCEL_EXIBE_LABELS_VALOR
            Case 6
                sNome = "Despesas"
                objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
                objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS 'EXCEL_EXIBE_LABELS_VALOR
            Case 7
                sNome = "Saldo"
                objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_LINE_MARKERS
                objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS 'EXCEL_EXIBE_LABELS_VALOR
            Case 8
                sNome = "Observação"
        
        End Select

        'Guarda o título da coluna em questão que será exibido na primeira célula
        objCelulas.vValor = sNome
        objCelulas.bFonteNegrito = True
         
         
        'Adiciona a célula à coleção de células
        objColunas.colCelulas.Add objCelulas

    Next
    
    For Each objRecebPagto In gobjProjeto.colRecebPagto
        For Each objRecebPagtoRegra In objRecebPagto.colRegras
            iParcela = 0
            For Each objParc In objRecebPagtoRegra.colParcelas
                Set objRecebPagtoRegraAux = New ClassPRJRecebPagtoRegras
            
                iParcela = iParcela + 1
                
                objRecebPagtoRegraAux.dtRegraValor = objRecebPagtoRegra.dtRegraValor
                objRecebPagtoRegraAux.iCondPagto = objRecebPagtoRegra.iCondPagto
                objRecebPagtoRegraAux.lNumIntDoc = objRecebPagtoRegra.lNumIntDoc
                objRecebPagtoRegraAux.lNumIntDocRecebPagto = objRecebPagtoRegra.lNumIntDocRecebPagto
                objRecebPagtoRegraAux.sObservacao = objRecebPagtoRegra.sObservacao
                objRecebPagtoRegraAux.sRegra = objRecebPagtoRegra.sRegra
                objRecebPagtoRegraAux.dValor = objRecebPagto.dValor
                objRecebPagtoRegraAux.dValorParcela = objParc.dValor
                objRecebPagtoRegraAux.iTipo = objRecebPagto.iTipo
                objRecebPagtoRegraAux.iTipo = objRecebPagto.iTipo
                objRecebPagtoRegraAux.iTipoNegativo = objRecebPagto.iTipo * -1 'Só para ordenação
                objRecebPagtoRegraAux.lNumero = objRecebPagto.lNumero
                objRecebPagtoRegraAux.dtDataVencimentoParcela = objParc.dtVencimento
                objRecebPagtoRegraAux.iParcela = iParcela
                colEntrada.Add objRecebPagtoRegraAux
            Next
        Next
    Next
    
    colCampos.Add "dtDataVencimentoParcela"
    colCampos.Add "iTipoNegativo"
    
    Call Ordena_Colecao(colEntrada, colRecebPagtoRegra, colCampos)
    
    For Each objRecebPagtoRegra In colRecebPagtoRegra
        For iIndice = 1 To 8

            'aponta para a coluna em questão
            Set objColunas = objPlanilha.colColunas(iIndice)
            
            'Instancia um novo objeto para armazenar os dados da célula
            Set objCelulas = New ClassCelulasExcel
            
            Select Case iIndice
            
                Case 1
                    If objRecebPagtoRegra.iTipo = PRJ_TIPO_PAGTO Then
                        objCelulas.vValor = "Pagamento " & CStr(objRecebPagtoRegra.lNumero)
                    Else
                        objCelulas.vValor = "Recebimento " & CStr(objRecebPagtoRegra.lNumero)
                    End If
                
                Case 2
                    objCelulas.vValor = objRecebPagtoRegra.iParcela
                    
                Case 3
                    objCelulas.vValor = objRecebPagtoRegra.sRegra
                
                Case 4
                    objCelulas.vValor = objRecebPagtoRegra.dtDataVencimentoParcela 'objRecebPagtoRegra.dtRegraValor
                  
                Case 5
                    If objRecebPagtoRegra.iTipo = PRJ_TIPO_RECEB Then
                        objCelulas.vValor = objRecebPagtoRegra.dValorParcela  'objRecebPagtoRegra.dValor * objRecebPagtoRegra.dPercentual
                        dSaldoAcumulado = dSaldoAcumulado + objRecebPagtoRegra.dValorParcela 'objRecebPagtoRegra.dValor * objRecebPagtoRegra.dPercentual
                    Else
                        objCelulas.vValor = 0
                    End If
                    objCelulas.sNumberFormat = "#,##0.00"
                
                Case 6
                    If objRecebPagtoRegra.iTipo <> PRJ_TIPO_RECEB Then
                        objCelulas.vValor = objRecebPagtoRegra.dValorParcela ' objRecebPagtoRegra.dValor * objRecebPagtoRegra.dPercentual
                        dSaldoAcumulado = dSaldoAcumulado - objRecebPagtoRegra.dValorParcela  'objRecebPagtoRegra.dValor * objRecebPagtoRegra.dPercentual
                    Else
                        objCelulas.vValor = 0
                    End If
                    objCelulas.sNumberFormat = "#,##0.00"
                
                Case 7
                    objCelulas.vValor = dSaldoAcumulado
                    objCelulas.sNumberFormat = "#,##0.00"
                
                Case 8
                    objCelulas.vValor = objRecebPagtoRegra.sObservacao
                    
            End Select
                               
            'adiciona a célula à coleção de células da coluna em questão
            objColunas.colCelulas.Add objCelulas
        Next
    Next
    
    'objPlanilha.lTipoGrafico = EXCEL_GRAFICO_3D_LINE
    
    'Informa ao excel o título do gráfico
    objPlanilha.sTituloGrafico = "Fluxo Financeiro do Projeto " & gobjProjeto.sNomeReduzido

    'Informa ao excel que o gráfico não deverá exibir legenda
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_DIREITA 'EXCEL_LEGENDA_NAO_EXIBE
    
    'Informa ao excel a posição dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW

    'Informa ao excel a orientação dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD

    'Informa ao excel que a plotagem do dados será por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS
    
    'Exporta os dados do objPlanilha para o Excel
    'lErro = CF("Excel_Gera_Planilha", objPlanilha)
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 187458
        
    Cronograma_Grafico_Excel = SUCESSO
    
    Exit Function
    
Erro_Cronograma_Grafico_Excel:
    
    Cronograma_Grafico_Excel = gErr
    
    Select Case gErr
        
        Case 187457, 187458
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187459)

    End Select

    Exit Function
    
End Function

Public Sub BotaoProposta_Click()

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_BotaoProposta_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 187460
    
    lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189315
    
    colSelecao.Add sProjeto

    Call Chama_Tela("PRJPropostasLista", colSelecao, objProposta, Nothing, "Projeto = ?")

    Exit Sub

Erro_BotaoProposta_Click:

    Select Case gErr

        Case 187460
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 189315

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187461)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer
Dim objProjeto As New ClassProjetos

On Error GoTo Erro_BotaoContrato_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 187460

    lErro = Projeto_Formata(Codigo.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189315
'
'    colSelecao.Add sProjeto
'
'    Call Chama_Tela("PRJContratosLista", colSelecao, objContrato, Nothing, "Projeto = ?")

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189697
    
    If lErro <> SUCESSO Then gError 189699

    objContrato.lNumIntDocPRJ = objProjeto.lNumIntDoc

    lErro = CF("PRJContratos_Le_Projeto", objContrato, False, False)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189698

    Call Chama_Tela("ContratoPRJ", objContrato)

    Exit Sub

Erro_BotaoContrato_Click:

    Select Case gErr

        Case 187460
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Codigo.SetFocus
            
        Case 189315, 189697, 189698
        
        Case 189699 'ERRO_PROJETO_NAO_CADASTRADO
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETO_NAO_CADASTRADO", gErr, sProjeto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187461)

    End Select

    Exit Sub
    
End Sub
Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
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

Public Function PRJ_Limpa_Tela_Projetos(ByVal objCT As Object) As Long
    PRJ_Limpa_Tela_Projetos = SUCESSO
End Function

Public Function PRJ_Traz_Projetos_Tela(ByVal objCT As Object, ByVal objProjeto As ClassProjetos) As Long
    PRJ_Traz_Projetos_Tela = SUCESSO
End Function

Private Function Carrega_Usuarios(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objUsuarios In colUsuarios
        objCombo.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190457)

    End Select

    Exit Function

End Function

Public Sub BotaoAnexos_Click()
   Call Chama_Tela_Modal("Anexos", gobjAnexos)
End Sub
