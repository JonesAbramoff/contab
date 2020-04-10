VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTLancamentoEstorno1"
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

Private objLancamento_Cabecalho1 As New ClassLancamento_Cabecalho
Private objLote1 As New ClassLote
Private objBrowseConfigura1 As AdmBrowseConfigura
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get Origem() As Object
     Set Origem = objUserControl.Controls("Origem")
End Property

Public Property Get Exercicio() As Object
     Set Exercicio = objUserControl.Controls("Exercicio")
End Property

Public Property Get Periodo() As Object
     Set Periodo = objUserControl.Controls("Periodo")
End Property

Public Property Get Documento() As Object
     Set Documento = objUserControl.Controls("Documento")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get LoteEstorno() As Object
     Set LoteEstorno = objUserControl.Controls("LoteEstorno")
End Property

Public Property Get UpDown1() As Object
     Set UpDown1 = objUserControl.Controls("UpDown1")
End Property

Public Property Get DataEstorno() As Object
     Set DataEstorno = objUserControl.Controls("DataEstorno")
End Property

Public Property Get DocumentoEstorno() As Object
     Set DocumentoEstorno = objUserControl.Controls("DocumentoEstorno")
End Property

Public Property Get PeriodoEstorno() As Object
     Set PeriodoEstorno = objUserControl.Controls("PeriodoEstorno")
End Property

Public Property Get ExercicioEstorno() As Object
     Set ExercicioEstorno = objUserControl.Controls("ExercicioEstorno")
End Property

Public Property Get LabelLoteEstorno() As Object
     Set LabelLoteEstorno = objUserControl.Controls("LabelLoteEstorno")
End Property

Public Property Get OrigemEstorno() As Object
     Set OrigemEstorno = objUserControl.Controls("OrigemEstorno")
End Property

Public Property Get LabelDocumentoEstorno() As Object
     Set LabelDocumentoEstorno = objUserControl.Controls("LabelDocumentoEstorno")
End Property

Public Property Get BotaoCancelar() As Object
     Set BotaoCancelar = objUserControl.Controls("BotaoCancelar")
End Property

Public Property Get BotaoOk() As Object
     Set BotaoOk = objUserControl.Controls("BotaoOk")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lDoc As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim dtData As Date

On Error GoTo Erro_BotaoProxNum_Click

    If Len(DataEstorno.ClipText) = 0 Then Error 55723

    'Obtém Periodo e Exercicio correspondentes à data
    dtData = CDate(DataEstorno.Text)

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 55724

    'Mostra número do próximo voucher(documento) disponível
    lErro = CF("Voucher_Automatico", giFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, MODULO_CONTABILIDADE, lDoc)
    If lErro <> SUCESSO Then Error 57515

    DocumentoEstorno.Text = CStr(lDoc)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 55723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNUM_DATA_NAO_PREENCHIDA", Err)
        
        Case 55724, 57515
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155979)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoCancelar_Click()

    objBrowseConfigura1.iTelaOK = CANCELA

    Unload Me

End Sub

Public Sub BotaoOk_Click()

Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Data dos Lançamentos de Estorno está preenchida
    If Len(DataEstorno.ClipText) = 0 Then gError 36923

    If Len(DocumentoEstorno.ClipText) = 0 Then gError 36925

    If CDate(DataEstorno.Text) < objLancamento_Cabecalho1.dtData Then gError 92050

'    If giTipoVersao = VERSAO_FULL Then
    
'        'Verifica se o Lote de Estorno está preenchido
'        If Len(LoteEstorno.ClipText) = 0 Then Error 36924

        'Preenche Objeto Lançamento_Cabeçalho
        objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
        objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
'        objLancamento_Cabecalho.iLote = CInt(LoteEstorno.ClipText)
        objLancamento_Cabecalho.dtData = CDate(DataEstorno.Text)
        objLancamento_Cabecalho.lDoc = CLng(DocumentoEstorno.Text)

        'grava o Estorno. objLancamento_Cabecalho1 contém o lote a ser extornado. objLancamento_Cabecalho contém algumas informacoes dos lançamentos de Estorno a serem criados
        lErro = CF("Lancamento_Grava_Estorno", objLancamento_Cabecalho, objLancamento_Cabecalho1)
        If lErro <> SUCESSO Then gError 36926
        
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Preenche Objeto Lançamento_Cabeçalho
'        objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
'        objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
'        objLancamento_Cabecalho.dtData = CDate(DataEstorno.Text)
'        objLancamento_Cabecalho.lDoc = CLng(DocumentoEstorno.Text)
'
'        'grava o Estorno. objLancamento_Cabecalho1 contém o lote a ser extornado. objLancamento_Cabecalho contém algumas informacoes dos lançamentos de Estorno a serem criados
'        lErro = CF("Lancamento_Grava_Estorno_Doc_Basico", objLancamento_Cabecalho, objLancamento_Cabecalho1)
'        If lErro <> SUCESSO Then gError 36926
'
'    End If
        
    objBrowseConfigura1.iTelaOK = OK

    GL_objMDIForm.MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 36923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            DataEstorno.SetFocus

        Case 36924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", gErr)
            LoteEstorno.SetFocus

        Case 36925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            DocumentoEstorno.SetFocus

        Case 36926

        Case 92050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_ESTORNO_MENOR_LANCAMENTO", gErr, DataEstorno.Text, objLancamento_Cabecalho1.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155980)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objLancamento_Cabecalho As ClassLancamento_Cabecalho, objBrowseConfigura As AdmBrowseConfigura) As Long

Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set objBrowseConfigura1 = objBrowseConfigura

    'guarda os dados de objLote para ser usado na gravação do Estorno
    objLancamento_Cabecalho1.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objLancamento_Cabecalho1.sOrigem = objLancamento_Cabecalho.sOrigem
    objLancamento_Cabecalho1.iExercicio = objLancamento_Cabecalho.iExercicio
    objLancamento_Cabecalho1.iPeriodoLan = objLancamento_Cabecalho.iPeriodoLan
    objLancamento_Cabecalho1.lDoc = objLancamento_Cabecalho.lDoc
    objLancamento_Cabecalho1.dtData = objLancamento_Cabecalho.dtData

    Lote.Caption = CStr(objLancamento_Cabecalho.iLote)

    Documento.Caption = CStr(objLancamento_Cabecalho.lDoc)

    lErro = CF("Exercicio_Le", objLancamento_Cabecalho.iExercicio, objExercicio)
    If lErro <> SUCESSO Then Error 36927

    Exercicio.Caption = objExercicio.sNomeExterno

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objPeriodo)
    If lErro <> SUCESSO Then Error 36928

    Periodo.Caption = objPeriodo.sNomeExterno

    Origem.Caption = gobjColOrigem.Descricao(objLancamento_Cabecalho.sOrigem)

    OrigemEstorno.Caption = gobjColOrigem.Descricao(MODULO_CONTABILIDADE)
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 36927, 36928

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155981)

    End Select

    Exit Function

End Function

Public Sub DataEstorno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEstorno)

End Sub

Public Sub DataEstorno_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim objLote As New ClassLote
Dim vbMsgRes As VbMsgBoxResult
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_DataEstorno_Validate

    If Len(DataEstorno.ClipText) > 0 Then

        lErro = Data_Critica(DataEstorno.Text)
        If lErro <> SUCESSO Then Error 36929

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36930

        'Verifica se Exercicio está fechado
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 36931

        'Exercicio não cadastrado
        If lErro = 10083 Then Error 36932

        If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36933

        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = objPeriodo.iExercicio
        objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
        objPeriodosFilial.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)


        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 36934

        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36935

        'checa se o lote pertence ao periodo em questão
        If Len(LoteEstorno.Text) > 0 Then

            objLote.iLote = CInt(LoteEstorno.Text)

            objLote.iFilialEmpresa = giFilialEmpresa
            objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
            objLote.iExercicio = objPeriodo.iExercicio
            objLote.iPeriodo = objPeriodo.iPeriodo

            'verifica se o lote  está atualizado
            lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
            If lErro <> SUCESSO Then Error 36936

            'Se é um lote que já foi contabilizado, não pode sofrer alteração
            If iLoteAtualizado = LOTE_ATUALIZADO Then Error 36937

            lErro = CF("LotePendente_Le", objLote)
            If lErro <> SUCESSO And lErro <> 5435 Then Error 36938

            'Se o lote não está cadastrado
            If lErro = 5435 Then Error 36939

            If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36940

        End If

        'Preenche campo de periodo
        PeriodoEstorno.Caption = objPeriodo.sNomeExterno

        ExercicioEstorno.Caption = objExercicio.sNomeExterno

    Else

        PeriodoEstorno.Caption = ""

        ExercicioEstorno.Caption = ""

    End If

    Exit Sub

Erro_DataEstorno_Validate:

    Cancel = True

    Select Case Err

        Case 36929, 36930, 36931, 36934

        Case 36932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 36935
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case 36937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, OrigemEstorno.Caption)

        Case 36936, 36938

        Case 36939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case 36940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155982)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoLote = Nothing
    Set objEventoLancamento = Nothing
    
End Sub

Public Sub DocumentoEstorno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DocumentoEstorno)

End Sub

Public Sub LoteEstorno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteEstorno)

End Sub

Public Sub LoteEstorno_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As New ClassLote
Dim sNomeExterno As String
Dim objExercicio As New ClassExercicio
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection

On Error GoTo Erro_LoteEstorno_Validate

    If Len(LoteEstorno.Text) > 0 And Len(DataEstorno.ClipText) > 0 Then

        objLote.iLote = CInt(LoteEstorno.Text)
        objLote.iFilialEmpresa = giFilialEmpresa
        objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36941

        objLote.iExercicio = objPeriodo.iExercicio
        objLote.iPeriodo = objPeriodo.iPeriodo

        'verifica se o lote  está atualizado
        lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
        If lErro <> SUCESSO Then Error 36942

        'Se é um lote que já foi contabilizado, não pode sofrer alteração
        If iLoteAtualizado = LOTE_ATUALIZADO Then Error 36943

        lErro = CF("LotePendente_Le", objLote)
        If lErro <> SUCESSO And lErro <> 5435 Then Error 36944

        'Se o lote não está cadastrado
        If lErro = 5435 Then Error 36945

        If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36946

    End If

    Exit Sub

Erro_LoteEstorno_Validate:

    Cancel = True

    Select Case Err

        Case 36941, 36942, 36944

        Case 36943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objPeriodo.iExercicio, objPeriodo.iPeriodo, OrigemEstorno.Caption)

        Case 36945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case 36946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155983)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim objExercicio As New ClassExercicio
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long

On Error GoTo Erro_Form_Load

    Set objEventoLote = New AdmEvento
    Set objEventoLancamento = New AdmEvento
    
    DataEstorno.Text = Format(gdtDataAtual, "dd/mm/yy")

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155984)

    End Select

    Exit Sub

End Sub

Public Sub LabelLoteEstorno_Click()

Dim objLote As New ClassLote
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelLoteEstorno_Click

    'Obtém Periodo e Exercicio correspondentes à data
    If Len(DataEstorno.ClipText) > 0 Then
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36949

    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If

    If Len(LoteEstorno.Text) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(LoteEstorno.Text)
    End If

    objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo

    Call Chama_Tela_Modal("LotePendenteListaModal", colSelecao, objLote, objEventoLote)

    Exit Sub

Erro_LabelLoteEstorno_Click:

    Select Case Err

        Case 36949

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155985)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As ClassLote
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1

    'Se estiver com a data preenchida ==> verificar se a data está dentro do periodo do lote
    If Len(DataEstorno.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36950

        'se o periodo/exercicio não corresponde ao periodo/exercicio do lote ==> troca a data
        If objPeriodo.iExercicio <> objLote.iExercicio Or objPeriodo.iPeriodo <> objLote.iPeriodo Then

            'move a data inicial do lote, exercicio e periodo para a tela
            lErro = Move_Data_Tela(objLote)
            If lErro <> SUCESSO Then Error 36951

        End If

    Else

        'se não estiver com a data preenchida
        'move a data inicial do lote, exercicio e periodo para a tela
        lErro = Move_Data_Tela(objLote)
        If lErro <> SUCESSO Then Error 36952

    End If

    LoteEstorno.Text = CStr(objLote.iLote)


    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case 36950, 36951, 36952  'Erro já tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155986)

    End Select

    Exit Sub

End Sub

Private Function Move_Data_Tela(objLote As ClassLote) As Long

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Move_Data_Tela

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 36953

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 36954

    'se o exercicio não estiver cadastrado
    If lErro = 10083 Then Error 36955

    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36956

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 36957

    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36958

    DataEstorno.Text = Format(objPeriodo.dtDataInicio, "dd/mm/yy")

    PeriodoEstorno.Caption = objPeriodo.sNomeExterno

    ExercicioEstorno.Caption = objExercicio.sNomeExterno

    Move_Data_Tela = SUCESSO

    Exit Function

Erro_Move_Data_Tela:

    Move_Data_Tela = Err

    Select Case Err

        Case 36953, 36954, 36957

        Case 36955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 36958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155987)

    End Select

    Exit Function

End Function

Public Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_UpDown1_DownClick

    If Len(Trim(DataEstorno.ClipText)) > 0 Then

        sData = DataEstorno.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 41688

        DataEstorno.Text = sData

    End If

    Call DataEstorno_Validate(bCancel)

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 41688

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155988)

    End Select

    Exit Sub

End Sub

Public Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_UpDown1_UpClick

    If Len(Trim(DataEstorno.ClipText)) > 0 Then

        sData = DataEstorno.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 41689

        DataEstorno.Text = sData

    End If

    Call DataEstorno_Validate(bCancel)

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 41689

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155989)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LANCAMENTO_EXTORNO1
    Set Form_Load_Ocx = Me
    Caption = "Estorno de Documento Contábil - Documento de Estorno"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LancamentoEstorno1"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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
        
        If Me.ActiveControl Is LoteEstorno Then
            Call LabelLoteEstorno_Click
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


