VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaTrataBloqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public objTela As Object
Private objUserControl As Object
Private gobjMapBloqGen As ClassMapeamentoBloqGen
Private giTipoTela As Integer
Private giGridBloqueioLinhasBloqueadas As Integer

'Grid Bloqueio
Public objGridBloqueio As AdmGrid
Public iGrid_TipoBloqueio_Col As Integer
Public iGrid_DataBloqueio_Col As Integer
Public iGrid_Usuario_Col As Integer
Public iGrid_Responsavel_Col As Integer
Public iGrid_DataLiberacao_Col As Integer
Public iGrid_ResponsavelLib_Col As Integer
Public iGrid_Observacao_Col As Integer
Public iGrid_SeqBloqueio_Col As Integer

Private Sub Class_Initialize()
    Set objGridBloqueio = New AdmGrid
End Sub

Private Sub Class_Terminate()
    Set objGridBloqueio = Nothing
    Set gobjMapBloqGen = Nothing
    Set objUserControl = Nothing
    Set objTela = Nothing
End Sub

Public Property Get TipoBloqueio() As Object
     Set TipoBloqueio = objUserControl.Controls("BLGTipoBloqueio")
End Property

Public Property Get DataBloqueio() As Object
     Set DataBloqueio = objUserControl.Controls("BLGDataBloqueio")
End Property

Public Property Get Usuario() As Object
     Set Usuario = objUserControl.Controls("BLGUsuario")
End Property

Public Property Get Responsavel() As Object
     Set Responsavel = objUserControl.Controls("BLGResponsavel")
End Property

Public Property Get DataLiberacao() As Object
     Set DataLiberacao = objUserControl.Controls("BLGDataLiberacao")
End Property

Public Property Get ResponsavelLib() As Object
     Set ResponsavelLib = objUserControl.Controls("BLGResponsavelLib")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("BLGObservacao")
End Property

Public Property Get SeqBloqueio() As Object
     Set SeqBloqueio = objUserControl.Controls("BLGSeqBloqueio")
End Property

Public Property Get GridBloqueio() As Object
     Set GridBloqueio = objUserControl.Controls("BLGGridBloqueio")
End Property

Public Function Inicializa(ByVal iTipoTela As Integer, ByVal objTelaGrid As Object, Optional ByVal bCT As Boolean = False) As Long

Dim lErro As Long
Dim objMapBloqGen As New ClassMapeamentoBloqGen

On Error GoTo Erro_Inicializa
    
    Set objTela = objTelaGrid
    
    If Not bCT Then
        Set objUserControl = objTelaGrid
    Else
        Set objUserControl = objTelaGrid.objUserControl
    End If
    
    giTipoTela = iTipoTela
    
    objMapBloqGen.iTipoTelaBloqueio = iTipoTela
    
    lErro = CF("MapeamentoBloqGen_Le", objMapBloqGen)
    If lErro <> SUCESSO Then gError 198400

    Set gobjMapBloqGen = objMapBloqGen
    
    giGridBloqueioLinhasBloqueadas = 0
    
    lErro = TiposDeBloqueios_Carrega(TipoBloqueio)
    If lErro <> SUCESSO Then gError 198435
    
    Set objGridBloqueio = New AdmGrid
    
    lErro = Inicializa_Grid_Bloqueio
    If lErro <> SUCESSO Then gError 198436
    
    Inicializa = SUCESSO

    Exit Function

Erro_Inicializa:

    Inicializa = gErr
    
    Select Case gErr
    
        Case 198400, 198435, 198436

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198401)

    End Select

    Exit Function
    
End Function

Public Function Inicializa_Grid_Bloqueio() As Long
'Inicializa o Grid

On Error GoTo Erro_Inicializa_Grid_Bloqueio

    'Form do Grid
    Set objGridBloqueio.objForm = objTela
    
    'Títulos das colunas
    objGridBloqueio.colColuna.Add (" ")
    objGridBloqueio.colColuna.Add ("Tipo")
    objGridBloqueio.colColuna.Add ("Data")
    objGridBloqueio.colColuna.Add ("Usuário")
    objGridBloqueio.colColuna.Add ("Responsável")
    objGridBloqueio.colColuna.Add ("Data Liberação")
    objGridBloqueio.colColuna.Add ("Resp. Liberação")
    objGridBloqueio.colColuna.Add ("Observação")
    objGridBloqueio.colColuna.Add ("Seq.")

    'Controles que participam do Grid
    objGridBloqueio.colCampo.Add (TipoBloqueio.Name)
    objGridBloqueio.colCampo.Add (DataBloqueio.Name)
    objGridBloqueio.colCampo.Add (Usuario.Name)
    objGridBloqueio.colCampo.Add (Responsavel.Name)
    objGridBloqueio.colCampo.Add (DataLiberacao.Name)
    objGridBloqueio.colCampo.Add (ResponsavelLib.Name)
    objGridBloqueio.colCampo.Add (Observacao.Name)
    objGridBloqueio.colCampo.Add (SeqBloqueio.Name)

    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_Usuario_Col = 3
    iGrid_Responsavel_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6
    iGrid_Observacao_Col = 7
    iGrid_SeqBloqueio_Col = 8

    'Grid do GridInterno
    objGridBloqueio.objGrid = GridBloqueio

    'Todas as linhas do grid
    objGridBloqueio.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    'Linhas visíveis do grid
    objGridBloqueio.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridBloqueio.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridBloqueio.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridBloqueio.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridBloqueio)

    Inicializa_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Bloqueio:

    Inicializa_Grid_Bloqueio = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198402)

    End Select

    Exit Function

End Function

Public Function Move_GridBloqueio_Memoria(ByVal colBloqueiosTela As Object) As Long
'Recolhe os dados do tab de Bloqueio

Dim lErro As Long
Dim objBloqueioGen As ClassBloqueioGen
Dim iIndice As Integer
Dim colBloqueios As New Collection

On Error GoTo Erro_Move_GridBloqueio_Memoria

    Set colBloqueios = ColecaoDef_Trans_Collection(colBloqueiosTela)

    If objGridBloqueio.iLinhasExistentes = 0 Then Exit Function

    'Recolhe os Bloqueios do Grid de Bloqueios
    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioGen = New ClassBloqueioGen

        objBloqueioGen.iTipoDeBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
        objBloqueioGen.dtData = StrParaDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col))
        objBloqueioGen.sCodUsuario = GridBloqueio.TextMatrix(iIndice, iGrid_Usuario_Col)
        objBloqueioGen.sResponsavel = GridBloqueio.TextMatrix(iIndice, iGrid_Responsavel_Col)
        objBloqueioGen.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
        objBloqueioGen.iSequencial = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col))
        objBloqueioGen.dtDataLib = StrParaDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))
        objBloqueioGen.sResponsavelLib = GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col)

        'Adiciona o bloqueio na coleção de bloqueios
        colBloqueios.Add objBloqueioGen

    Next
    
    Call Collection_Trans_ColecaoDef(colBloqueios, colBloqueiosTela)

    Move_GridBloqueio_Memoria = SUCESSO

    Exit Function

Erro_Move_GridBloqueio_Memoria:

    Move_GridBloqueio_Memoria = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198403)

    End Select

    Exit Function

End Function

Public Sub TipoBloqueio_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataBloqueio_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub DataBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub DataBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = DataBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Observacao_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Usuario_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Usuario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub Usuario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub Usuario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = Usuario
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Responsavel_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Responsavel_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub Responsavel_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = Responsavel
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, objTela.iAlterado)
    End If

End Sub

Public Sub GridBloqueio_GotFocus()
    Call Grid_Recebe_Foco(objGridBloqueio)
End Sub

Public Sub GridBloqueio_EnterCell()
    Call Grid_Entrada_Celula(objGridBloqueio, objTela.iAlterado)
End Sub

Public Sub GridBloqueio_LeaveCell()
    Call objTela.Saida_Celula(objGridBloqueio)
End Sub

Public Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)

End Sub

Public Sub GridBloqueio_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, objTela.iAlterado)
    End If

End Sub

Public Sub GridBloqueio_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridBloqueio)
End Sub

Public Sub GridBloqueio_RowColChange()
    Call Grid_RowColChange(objGridBloqueio)
End Sub

Public Sub GridBloqueio_Scroll()
    Call Grid_Scroll(objGridBloqueio)
End Sub

Public Function Saida_Celula(ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'aquii está devolvendo erro em vez de sucesso
'    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
'    If lErro = SUCESSO Then
   
        'Verifica qual a coluna atual do Grid
        Select Case objGridBloqueio.objGrid.Col
    
            Case iGrid_TipoBloqueio_Col
                lErro = Saida_Celula_TipoBloqueio(objGridInt)
                If lErro <> SUCESSO Then gError 198404
    
            Case iGrid_Responsavel_Col
                lErro = Saida_Celula_Responsavel(objGridInt)
                If lErro <> SUCESSO Then gError 198405
    
            Case iGrid_Observacao_Col
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 198406
    
        End Select
        
'        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'        If lErro <> SUCESSO Then gError 198406

'    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 198404 To 198406

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198407)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoBloqueio.Text <> TipoBloqueio.List(TipoBloqueio.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 198408

            'Não foi encontrado
            If lErro = 25085 Then gError 198409
            If lErro = 25086 Then gError 198410

        End If

        'Verifica se já existe um bloqueio, não liberado, do mesmo tipo no grid.
        For iIndice = 1 To objGridInt.iLinhasExistentes

            If iIndice <> GridBloqueio.Row Then
                iCodigo = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))

                If iCodigo = TipoBloqueio.ItemData(TipoBloqueio.ListIndex) Then
                    If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) = 0 Then gError 198411
                End If
            End If

        Next

        'Acrescenta uma linha no Grid se for o caso
        If GridBloqueio.Row - GridBloqueio.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

            'Coloca data e usuário
            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_DataBloqueio_Col) = Format(gdtDataHoje, "dd/mm/yy")
            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_Usuario_Col) = Format(gsUsuario)

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198412

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = gErr

    Select Case gErr

        Case 198408, 198412
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 198409
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 198410
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 198411
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_REPETIDO", gErr, Codigo_Extrai(TipoBloqueio.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198413)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Responsavel(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Responsável BL que está deixando de serr a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Responsavel

    Set objGridInt.objControle = Responsavel

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198414

    Saida_Celula_Responsavel = SUCESSO

    Exit Function

Erro_Saida_Celula_Responsavel:

    Saida_Celula_Responsavel = gErr

    Select Case gErr

        Case 198414
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198415)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198416

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 198416
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198417)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoLibBloq_Click(ByVal objControleChave As Object)

Dim lErro As Long
Dim objDocBloq As Object
Dim colBloqueios As Collection

On Error GoTo Erro_BotaoLibBloq_Click

    Set objDocBloq = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)

    'Verifica se o Código foi preenchido
    If Len(Trim(objControleChave.ClipText)) > 0 Then

        'Passa os dados do Bloqueio para o Obj
        If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
            objDocBloq.iFilialEmpresa = giFilialEmpresa
        End If
        
        Call CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbLet, StrParaDbl(objControleChave.Text))

        'Lê o Documento
        lErro = CF(gobjMapBloqGen.sNomeFuncLeDoc, objDocBloq)
        If lErro <> SUCESSO Then gError 198418
        
    End If

    Call Chama_Tela_Modal("LiberaBloqueioGen", gobjMapBloqGen.iTipoTelaBloqueio, objDocBloq)

    'Le os Bloqueios
    lErro = CF("BloqueiosGen_Le", gobjMapBloqGen, objDocBloq)
    If lErro <> SUCESSO Then gError 198419
    
    Set colBloqueios = ColecaoDef_Trans_Collection(CallByName(objDocBloq, gobjMapBloqGen.sNomeColecaoBloqDoc, VbGet))

    'Recarrega o tab de Bloqueios
    lErro = Carrega_GridBloqueio(objDocBloq)
    If lErro <> SUCESSO Then gError 198420

    Exit Sub

Erro_BotaoLibBloq_Click:

    Select Case gErr

        Case 198418 To 198420

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198421)

    End Select

    Exit Sub

End Sub

Public Function Carrega_GridBloqueio(ByVal objDocBloq As Object) As Long
'Coloca os dados do Tab de Bloqueio na tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objBloqueioGen As ClassBloqueioGen
Dim objTipoBloqueio As New ClassTipoDeBloqueio
Dim colBloqueios As Collection
Dim colBloqueiosDoc As Variant
Dim objBloq As Variant

On Error GoTo Erro_Carrega_GridBloqueio

    Set colBloqueiosDoc = CallByName(objDocBloq, gobjMapBloqGen.sNomeColecaoBloqDoc, VbGet)

    For Each objBloq In colBloqueiosDoc
        colBloqueiosDoc.Remove 1
    Next
    
    lErro = CF("BloqueiosGen_le", gobjMapBloqGen, objDocBloq)
    If lErro <> SUCESSO Then gError 198534

    Set colBloqueios = ColecaoDef_Trans_Collection(colBloqueiosDoc)
    
    'Limpa o Grid de Bloqueios antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioGen In colBloqueios

        iIndice = iIndice + 1
        
        'Coloca o bloqieo no Grid de bloqueios
        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioGen.sNomeTipoDeBloqueio
        GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioGen.dtData, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iIndice, iGrid_Usuario_Col) = objBloqueioGen.sCodUsuario
        GridBloqueio.TextMatrix(iIndice, iGrid_Responsavel_Col) = objBloqueioGen.sResponsavel
        If objBloqueioGen.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioGen.dtDataLib, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioGen.sResponsavelLib
        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioGen.sObservacao
        GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col) = CStr(objBloqueioGen.iSequencial)

    Next
    
    'Guarda o número de bloqueios trazidos do BD
    giGridBloqueioLinhasBloqueadas = colBloqueios.Count

    objGridBloqueio.iLinhasExistentes = iIndice

    Carrega_GridBloqueio = SUCESSO

    Exit Function

Erro_Carrega_GridBloqueio:

    Carrega_GridBloqueio = gErr

    Select Case gErr

        Case 198422

        Case 198423
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", gErr, objTipoBloqueio.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198424)

    End Select

End Function

Public Function Valida_Grid_Bloqueio() As Long
'Validação de GridBloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim bBloqueioTotal As Boolean
Dim iTipoBloqueio As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataLiberacao As Date

On Error GoTo Erro_Valida_Grid_Bloqueio

    bBloqueioTotal = False

    If objGridBloqueio.iLinhasExistentes > giGridBloqueioLinhasBloqueadas Then

        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

            'Testa se preencheu tipo de bloqueio
            If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))) = 0 Then gError 198425

        Next

    End If

    If objGridBloqueio.iLinhasExistentes > 0 Then

        'Testa se existe bloqueio total não liberado
        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

            iTipoBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
            dtDataLiberacao = StrParaDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))

            If iTipoBloqueio = BLOQUEIO_TOTAL And dtDataLiberacao = DATA_NULA Then bBloqueioTotal = True

        Next

        If bBloqueioTotal Then

            'Pergunta se deseja continuar com bloqueio total existente
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BLOQUEIO_TOTAL")
            If vbMsgRes = vbNo Then gError 198426

        End If

    End If

    Valida_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Valida_Grid_Bloqueio:

    Valida_Grid_Bloqueio = gErr

    Select Case gErr

        Case 198425
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", gErr, iIndice)

        Case 198426

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198427)

    End Select

    Exit Function

End Function

Public Function Limpa() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa

    Call Grid_Limpa(objGridBloqueio)
    
    giGridBloqueioLinhasBloqueadas = 0
    
    Limpa = SUCESSO

    Exit Function

Erro_Limpa:

    Limpa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198428)

    End Select

    Exit Function
    
End Function

Private Function TiposDeBloqueios_Carrega(ByVal objCombo As Object) As Long
'Carrega a lista de tipos de bloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDeBloqueio As New Collection
Dim objTipoDeBloqueio As ClassTiposDeBloqueioGen

On Error GoTo Erro_TiposDeBloqueios_Carrega

    'Le todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", gobjMapBloqGen.iTipoTelaBloqueio, colTipoDeBloqueio)
    If lErro <> SUCESSO And lErro <> 29168 Then gError 198381

    'Preenche ListaTipos
    For Each objTipoDeBloqueio In colTipoDeBloqueio
        If objTipoDeBloqueio.iNaoApareceTelaLib = DESMARCADO Then
            objCombo.AddItem objTipoDeBloqueio.iCodigo & SEPARADOR & objTipoDeBloqueio.sNomeReduzido
            objCombo.ItemData(objCombo.NewIndex) = objTipoDeBloqueio.iCodigo
        End If
    Next

    TiposDeBloqueios_Carrega = SUCESSO

    Exit Function

Erro_TiposDeBloqueios_Carrega:

    TiposDeBloqueios_Carrega = gErr

    Select Case gErr

        Case 198381

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198382)

    End Select

    Exit Function

End Function

Public Function Rotina_Grid_Enabled(iLinha As Integer, objControl As Object, iCaminho As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enabled
    
    Select Case objControl.Name

        Case TipoBloqueio.Name, Responsavel.Name
            'Se for algum Bloqueio previamente cadastrado, desabilita o controle
            If iLinha <= giGridBloqueioLinhasBloqueadas Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select
            
    Rotina_Grid_Enabled = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enabled:

    Rotina_Grid_Enabled = gErr
    
    Select Case gErr
    
        Case 198400, 198435, 198436

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198401)

    End Select

    Exit Function
    
End Function

Public Function Trata_Bloqueios_Gravacao(ByVal objDocGravado As Object) As Long

Dim lErro As Long
Dim colBloqueios As New Collection
Dim bExibirBloqs As Boolean

On Error GoTo Erro_Trata_Bloqueios_Gravacao

    bExibirBloqs = False

    Call Doc_ExibirBloqs(objDocGravado, colBloqueios, bExibirBloqs)
    If bExibirBloqs Then

        lErro = Chama_Tela("BloqueiosGen_ConsultaTodos", gobjMapBloqGen, objDocGravado)
        If lErro <> SUCESSO Then gError 198438

    End If
            
    Trata_Bloqueios_Gravacao = SUCESSO

    Exit Function

Erro_Trata_Bloqueios_Gravacao:

    Trata_Bloqueios_Gravacao = gErr
    
    Select Case gErr
    
        Case 198438

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198440)

    End Select

    Exit Function
    
End Function

Private Sub Doc_ExibirBloqs(ByVal objDocBloq As Object, ByVal colBloqueios As Collection, bExibirBloqs As Boolean)
'retorna em bExibirBloqsPV se deve abrir a tela que mostra os bloqueios de pv

Dim lErro As Long
Dim objBloq As Variant
Dim objBloqAux As ClassBloqueioPV
Dim bCriado As Boolean
Dim colBloqueiosDoc As Variant

On Error GoTo Erro_Doc_ExibirBloqs

    Set colBloqueiosDoc = CallByName(objDocBloq, gobjMapBloqGen.sNomeColecaoBloqDoc, VbGet)

    For Each objBloq In colBloqueiosDoc
        colBloqueiosDoc.Remove 1
    Next
    
    lErro = CF("BloqueiosGen_le", gobjMapBloqGen, objDocBloq)
    If lErro <> SUCESSO Then gError 198534

    'verifica se algum bloqueio automatico foi criado durante a gravacao
    bCriado = False

    For Each objBloq In colBloqueiosDoc

        If objBloq.dtDataLib = DATA_NULA Then

            bCriado = True

            For Each objBloqAux In colBloqueios

                If objBloqAux.iTipoDeBloqueio = objBloq.iTipoDeBloqueio Then

                    bCriado = False
                    Exit For

                End If

            Next

            If bCriado Then Exit For

        End If

    Next

    bExibirBloqs = bCriado

    Exit Sub

Erro_Doc_ExibirBloqs:

    Select Case gErr
    
        Case 198534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198439)

    End Select

    Exit Sub

End Sub
