VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

Dim m_Caption As String
Event Unload()

Const ESTCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const ESTCONFIG_GERA_LOTE_AUTOMATICO = 1
Const ESTCONFIG_INVENTARIOCODBARRAAUTO = 2

'Inserido por Jorge Specian - 06/04/2005
'---------------------------------------
'Grid de Prioridades
Dim objGridPrioridades As AdmGrid
Dim iGrid_Prioridade_Col As Integer

Private WithEvents objEventoClasseUM As AdmEvento
Attribute objEventoClasseUM.VB_VarHelpID = -1
'----------------------------------------

Dim iAlterado As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get BloqueioCTB() As Object
     Set BloqueioCTB = objUserControl.Controls("BloqueioCTB")
End Property

Public Property Get DataBloqLimite() As Object
     Set DataBloqLimite = objUserControl.Controls("DataBloqLimite")
End Property

Public Property Get UpDownDataBloqLimite() As Object
     Set UpDownDataBloqLimite = objUserControl.Controls("UpDownDataBloqLimite")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get ListaConfigura() As Object
     Set ListaConfigura = objUserControl.Controls("ListaConfigura")
End Property

Public Property Get IntervaloProducao() As Object
     Set IntervaloProducao = objUserControl.Controls("IntervaloProducao")
End Property

Public Property Get IncluiFrete() As Object
     Set IncluiFrete = objUserControl.Controls("IncluiFrete")
End Property

Public Property Get AceitaQtdNegativa() As Object
     Set AceitaQtdNegativa = objUserControl.Controls("AceitaQtdNegativa")
End Property


'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155639)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim lErro As Long
Dim iIntervalo As Integer

On Error GoTo Erro_Form_Load
    
    'Checa Aglutina lançamentos por dia
    If gobjEST.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
        ListaConfigura.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
    Else
        ListaConfigura.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
    End If
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Checa "Gera Lote Automatico"
        If gobjEST.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
            ListaConfigura.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = True
        Else
            ListaConfigura.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = False
        End If
        
'    End If
    
    If gobjEST.iInventarioCodBarrAuto = INVENTARIOCODBARRAAUTO_SIM Then
        ListaConfigura.Selected(ESTCONFIG_INVENTARIOCODBARRAAUTO) = True
    Else
        ListaConfigura.Selected(ESTCONFIG_INVENTARIOCODBARRAAUTO) = False
    End If
    
    'Verifica se inclui frete e outras despesas no cálculo do custo
    If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
        IncluiFrete.Value = vbChecked
    Else
        IncluiFrete.Value = vbUnchecked
    End If
    
    If gobjMAT.iAceitaEstoqueNegativo = MARCADO Then
        AceitaQtdNegativa.Value = vbChecked
    Else
        AceitaQtdNegativa.Value = vbUnchecked
    End If
    
    If gobjMAT.iMovEstVerificaBloqData = MARCADO Then
        BloqueioCTB.Value = vbChecked
    Else
        BloqueioCTB.Value = vbUnchecked
    End If
    
    If gobjMAT.dtMovEstBloqDataLimite <> DATA_NULA Then Call DateParaMasked(DataBloqLimite, gobjMAT.dtMovEstBloqDataLimite)
    
    'Inserido por Jorge Specian - 30/06/2005
    '---------------------------------------
    
    'Grid Prioridades
    Set objGridPrioridades = New AdmGrid
    
    'tela em questão
    Set objGridPrioridades.objForm = Me
    
    lErro = Inicializa_GridPrioridades(objGridPrioridades)
    If lErro <> SUCESSO Then gError 134639
    
    If gobjEST.dtDataInicioMRP <> DATA_NULA Then
        
        DataInicio.Text = Format(gobjEST.dtDataInicioMRP, "dd/mm/yy")
    
        GridPrioridades.TextMatrix(gobjEST.iPrioridadeMaquina, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_MAQUINAS
        GridPrioridades.TextMatrix(gobjEST.iPrioridadeProduto, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS
        GridPrioridades.TextMatrix(gobjEST.iPrioridadeProdutoAnc, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS_ANC
    
        objGridPrioridades.iLinhasExistentes = NUM_PRIORIDADES_SELECAO
    
    End If
    
    If gobjEST.iClasseUMTempo <> 0 Then
    
        ClasseUMTempo.Text = CStr(gobjEST.iClasseUMTempo)
        
    End If
        
    If gobjEST.iGeraReqCompraEmLote = GERA_REQCOMPRA_EM_LOTE Then
        GeraReqCompraEmLote.Value = vbChecked
    Else
        GeraReqCompraEmLote.Value = vbUnchecked
    End If
    
    Set objEventoClasseUM = New AdmEvento

    '---------------------------------------
    
    lErro = CF("IntervaloProducao_Le", iIntervalo)
    If lErro <> SUCESSO Then gError 76006
        
    IntervaloProducao.Text = iIntervalo
     
    lErro = CF2(Me, "CTConfig_Form_Load")
    If lErro <> SUCESSO Then gError 76006
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 76006, 134639
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155640)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 57575
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 57575 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155641)

    End Select

    Exit Sub
    
End Sub
    
Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sIntervaloProducao As String
Dim iIndice As Integer   'Inserido por Jorge Specian

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Move os Dados para Memoria
    If ListaConfigura.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjEST.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjEST.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
'    If giTipoVersao = VERSAO_FULL Then
    
        If ListaConfigura.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = True Then
            gobjEST.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
        Else
            gobjEST.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
    
'    End If
    
    If ListaConfigura.Selected(ESTCONFIG_INVENTARIOCODBARRAAUTO) = True Then
        gobjEST.iInventarioCodBarrAuto = INVENTARIOCODBARRAAUTO_SIM
    Else
        gobjEST.iInventarioCodBarrAuto = INVENTARIOCODBARRAAUTO_NAO
    End If
    
    sIntervaloProducao = IntervaloProducao.Text
    
    If IncluiFrete.Value = vbChecked Then
        gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO
    Else
        gobjEST.iIncluiFreteNoCusto = NAO_INCLUI_FRETE_CUSTO
    End If
    
    If AceitaQtdNegativa.Value = vbChecked Then
        gobjMAT.iAceitaEstoqueNegativo = MARCADO
    Else
        gobjMAT.iAceitaEstoqueNegativo = DESMARCADO
    End If
    
    'Inserido por Jorge Specian - 30/06/2005
    '---------------------------------------
    
    'Verifica se a Data Inicial Operações MRP está preenchida
    If Len(DataInicio.ClipText) > 0 Then
    
        If Len(Trim(ClasseUMTempo.Text)) = 0 Then gError 134614
        
        lErro = Valor_Positivo_Critica(ClasseUMTempo.Text)
        If lErro <> SUCESSO Then gError 134615
    
        gobjEST.dtDataInicioMRP = StrParaDate(DataInicio.Text)
        gobjEST.iClasseUMTempo = StrParaInt(ClasseUMTempo.Text)
        
        For iIndice = 1 To objGridPrioridades.iLinhasExistentes
        
            Select Case GridPrioridades.TextMatrix(iIndice, iGrid_Prioridade_Col)
            
                Case Is = STRING_GRIDPRIORIDADE_MAQUINAS
                    gobjEST.iPrioridadeMaquina = iIndice
            
                Case Is = STRING_GRIDPRIORIDADE_PRODUTOS
                    gobjEST.iPrioridadeProduto = iIndice
            
                Case Is = STRING_GRIDPRIORIDADE_PRODUTOS_ANC
                    gobjEST.iPrioridadeProdutoAnc = iIndice
            
            End Select
        
        Next iIndice
            
    Else
    
        gobjEST.dtDataInicioMRP = CDate(DATA_NULA)
        
    End If
    
    If GeraReqCompraEmLote.Value = vbChecked Then
        gobjEST.iGeraReqCompraEmLote = GERA_REQCOMPRA_EM_LOTE
    Else
        gobjEST.iGeraReqCompraEmLote = NAO_GERA_REQCOMPRA_EM_LOTE
    End If

    lErro = CF2(Me, "CTConfig_Gravar_Registro")
    If lErro <> SUCESSO Then gError 76007
    
    '---------------------------------------
    
    If BloqueioCTB.Value = vbChecked Then
        gobjMAT.iMovEstVerificaBloqData = MARCADO
    Else
        gobjMAT.iMovEstVerificaBloqData = DESMARCADO
    End If
    
    gobjMAT.dtMovEstBloqDataLimite = StrParaDate(DataBloqLimite.Text)
    
    'Grava na Tabela  "ESTConfig" as Configurações
    lErro = gobjEST.Gravar()
    If lErro <> SUCESSO Then gError 57576

    lErro = CF("MatConfig_Atualiza_IntervaloProducao", sIntervaloProducao)
    If lErro <> SUCESSO Then gError 76007
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 57576, 76007, 134613, 134615
            'erros tratados nas Rotinas Chamadas
        
        Case 134614
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSEUM_NAO_PREENCHIDO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155642)
            
    End Select

    Exit Function
    
End Function

Public Sub ListaConfigura_ItemCheck(Item As Integer)
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub IntervaloProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub IncluiFrete_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub
Public Sub AceitaQtdNegativa_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'obs.: nao apague pois é chamado da tela (usercontrol)
Public Sub Form_UnLoad(Cancel As Integer)

    Set gobjInfoUsu = Nothing

    'Incluido por Jorge Specian
    '---------------------------
    Set objGridPrioridades = Nothing
    Set objEventoClasseUM = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURACAOEST
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraEST"
    
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

'Inserido por Jorge Specian - 06/04/2005
'---------------------------------------
Private Function Inicializa_GridPrioridades(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Prioridades")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Prioridade.Name)

    'Colunas do Grid
    iGrid_Prioridade_Col = 1

    objGrid.objGrid = GridPrioridades

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 4

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridPrioridades.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iProibidoIncluirNoMeioGrid = 0
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridPrioridades = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
                
    'Grid Prioridades
    If objControl.Name = "Prioridade" Then
        
        objControl.Enabled = False
        
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155643)

    End Select

    Exit Sub

End Sub

Public Property Get GridPrioridades() As Object
     Set GridPrioridades = objUserControl.Controls("GridPrioridades")
End Property

Public Sub GridPrioridades_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPrioridades, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPrioridades, iAlterado)
    End If

End Sub

Public Sub GridPrioridades_EnterCell()

    Call Grid_Entrada_Celula(objGridPrioridades, iAlterado)

End Sub

Public Sub GridPrioridades_GotFocus()

    Call Grid_Recebe_Foco(objGridPrioridades)

End Sub

Public Sub GridPrioridades_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPrioridades, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPrioridades, iAlterado)
    End If

End Sub

Public Sub GridPrioridades_LeaveCell()

    Call Saida_Celula(objGridPrioridades)

End Sub

Public Sub GridPrioridades_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPrioridades)
    
End Sub

Public Sub GridPrioridades_RowColChange()

    Call Grid_RowColChange(objGridPrioridades)

End Sub

Public Sub GridPrioridades_Scroll()

    Call Grid_Scroll(objGridPrioridades)

End Sub

Public Property Get Prioridade() As Object
     Set Prioridade = objUserControl.Controls("Prioridade")
End Property

Public Sub Prioridade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Public Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrioridades)

End Sub

Public Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrioridades)

End Sub

Public Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridPrioridades.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGridPrioridades)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    Set objGridInt.objControle = Prioridade
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134392

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
        
        Case 134392
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155644)

    End Select

    Exit Function

End Function

Public Property Get UpDownPrioridade() As Object
     Set UpDownPrioridade = objUserControl.Controls("UpDownPrioridade")
End Property

Public Sub DownPrioridade_Click()

Dim lErro As Long
Dim iLinhaAtual As Integer
Dim iLinhaDebaixo As Integer
Dim sPrioridade1 As String
Dim sPrioridade2 As String

On Error GoTo Erro_DownPrioridade_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPrioridades.Row = 0 Then gError 134611
    
    If GridPrioridades.Row = NUM_PRIORIDADES_SELECAO Then Exit Sub
    
    iLinhaAtual = GridPrioridades.Row
    
    iLinhaDebaixo = iLinhaAtual + 1

    sPrioridade1 = GridPrioridades.TextMatrix(iLinhaAtual, iGrid_Prioridade_Col)
    sPrioridade2 = GridPrioridades.TextMatrix(iLinhaDebaixo, iGrid_Prioridade_Col)
    
    GridPrioridades.TextMatrix(iLinhaDebaixo, iGrid_Prioridade_Col) = sPrioridade1
    GridPrioridades.TextMatrix(iLinhaAtual, iGrid_Prioridade_Col) = sPrioridade2
    
    GridPrioridades.Row = iLinhaDebaixo
    
    Exit Sub

Erro_DownPrioridade_Click:

    Select Case gErr

        Case 134611
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155645)

    End Select

    Exit Sub

End Sub

Public Sub UpPrioridade_Click()

Dim lErro As Long
Dim iLinhaAtual As Integer
Dim iLinhaDeCima As Integer
Dim sPrioridade1 As String
Dim sPrioridade2 As String

On Error GoTo Erro_UpPrioridade_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPrioridades.Row = 0 Then gError 134612
    
    If GridPrioridades.Row = 1 Then Exit Sub
    
    iLinhaAtual = GridPrioridades.Row
    
    iLinhaDeCima = iLinhaAtual - 1

    sPrioridade1 = GridPrioridades.TextMatrix(iLinhaAtual, iGrid_Prioridade_Col)
    sPrioridade2 = GridPrioridades.TextMatrix(iLinhaDeCima, iGrid_Prioridade_Col)
    
    GridPrioridades.TextMatrix(iLinhaDeCima, iGrid_Prioridade_Col) = sPrioridade1
    GridPrioridades.TextMatrix(iLinhaAtual, iGrid_Prioridade_Col) = sPrioridade2

    GridPrioridades.Row = iLinhaDeCima

    Exit Sub

Erro_UpPrioridade_Click:

    Select Case gErr

        Case 134612
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155646)

    End Select

    Exit Sub

End Sub

Public Property Get ClasseUMTempo() As Object
     Set ClasseUMTempo = objUserControl.Controls("ClasseUMTempo")
End Property

Public Property Get DataInicio() As Object
     Set DataInicio = objUserControl.Controls("DataInicio")
End Property

Public Property Get UpDownDataInicio() As Object
     Set UpDownDataInicio = objUserControl.Controls("UpDownDataInicio")
End Property

Public Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    DataInicio.SetFocus

    If Len(DataInicio.ClipText) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137690

        DataInicio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 137690

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155647)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    DataInicio.SetFocus

    If Len(Trim(DataInicio.ClipText)) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137691

        DataInicio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 137691

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155648)

    End Select

    Exit Sub

End Sub

Public Property Get GeraReqCompraEmLote() As Object
     Set GeraReqCompraEmLote = objUserControl.Controls("GeraReqCompraEmLote")
End Property

Public Sub GeraReqCompraEmLote_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Property Get LabelClasseUMTempo() As Object
     Set LabelClasseUMTempo = objUserControl.Controls("LabelClasseUMTempo")
End Property

Public Sub LabelClasseUMTempo_Click()

Dim objClasseUM As ClassUnidadeDeMedida
Dim colSelecao As Collection

    Call Chama_Tela("ClasseUMLista", colSelecao, objClasseUM, objEventoClasseUM)

End Sub

Public Sub objEventoClasseUM_evSelecao(obj1 As Object)

Dim objClasseUM As New ClassClasseUM
Dim bCancel As Boolean

    Set objClasseUM = obj1

    'Preenche Text da ClasseUM
    ClasseUMTempo.Text = CStr(objClasseUM.iClasse)

    Me.Show

End Sub

Public Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    If Len(Trim(DataInicio.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError 137934
        
        If gobjEST.iClasseUMTempo > 0 Then ClasseUMTempo.Text = CStr(gobjEST.iClasseUMTempo)
    
        If gobjEST.iPrioridadeMaquina = 0 Or gobjEST.iPrioridadeProduto = 0 Or gobjEST.iPrioridadeProdutoAnc Then
                
            GridPrioridades.TextMatrix(1, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_MAQUINAS
            GridPrioridades.TextMatrix(2, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS
            GridPrioridades.TextMatrix(3, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS_ANC
        
        Else
        
            GridPrioridades.TextMatrix(gobjEST.iPrioridadeMaquina, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_MAQUINAS
            GridPrioridades.TextMatrix(gobjEST.iPrioridadeProduto, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS
            GridPrioridades.TextMatrix(gobjEST.iPrioridadeProdutoAnc, iGrid_Prioridade_Col) = STRING_GRIDPRIORIDADE_PRODUTOS_ANC
        
        End If
        
        objGridPrioridades.iLinhasExistentes = NUM_PRIORIDADES_SELECAO

    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137934
            'erro tratado na rotina chamada
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155649)

    End Select

    Exit Sub
    
End Sub

'---------------------------------------
Public Function CTConfig_Form_Load(ByVal objCT As Object) As Long
    CTConfig_Form_Load = SUCESSO
End Function

Public Function CTConfig_Gravar_Registro(ByVal objCT As Object) As Long
    CTConfig_Gravar_Registro = SUCESSO
End Function

Public Sub UpDownDataBloqLimite_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_DownClick

    DataBloqLimite.SetFocus

    If Len(DataBloqLimite.ClipText) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155647)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataBloqLimite_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_UpClick

    DataBloqLimite.SetFocus

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155648)

    End Select

    Exit Sub

End Sub

Public Sub DataBloqLimite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBloqLimite_Validate

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataBloqLimite.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DataBloqLimite_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            'erro tratado na rotina chamada
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155649)

    End Select

    Exit Sub
    
End Sub

Public Sub BloqueioCTB_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataBloqLimite_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
