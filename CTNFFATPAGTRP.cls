VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFATPAGTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim gobjCT As Object
Private WithEvents objEventoAporte As AdmEvento
Attribute objEventoAporte.VB_VarHelpID = -1

Public Property Get Aporte(ByVal objCT As Object) As Object
     Set Aporte = objCT.objUserControl.Controls("Aporte")
End Property

Public Property Get TipoPagto(ByVal objCT As Object) As Object
     Set TipoPagto = objCT.objUserControl.Controls("TipoPagto")
End Property

Public Property Get Pagto(ByVal objCT As Object) As Object
     Set Pagto = objCT.objUserControl.Controls("Pagto")
End Property

Public Sub Aporte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Aporte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objAporte As New ClassTRPAportes

On Error GoTo Erro_APorte_Validate

    'Se a data APorte está preenchida
    If Len(Trim(Aporte(objCT).Text)) > 0 Then

        objAporte.lCodigo = StrParaLong(Aporte(objCT).Text)
        
        'Lê o TRPAporte que está sendo Passado
        lErro = CF("TRPAportes_Le", objAporte)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192349
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 192330
        
        If TipoPagto(objCT).ListIndex <> -1 Then
            lErro = Trata_Troca_Aporte_TipoPagto(objCT, objAporte.lCodigo, TipoPagto(objCT).ItemData(TipoPagto(objCT).ListIndex))
            If lErro <> SUCESSO Then gError 192355
        End If

    Else
    
        Pagto(objCT).Clear

    End If
    
    Set objCT.gobjinfousu.gobjAporte = objAporte

    Exit Sub

Erro_APorte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 192330  'ERRO_TRPAPORTES_NAO_CADASTRADO
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, objAporte.lCodigo)

        Case 192349, 192355

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192331)

    End Select

    Exit Sub

End Sub

Public Sub TipoPagto_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
    If TipoPagto(objCT).ListIndex <> -1 Then
        Call Trata_Troca_Aporte_TipoPagto(objCT, objCT.gobjinfousu.gobjAporte.lCodigo, TipoPagto(objCT).ItemData(TipoPagto(objCT).ListIndex))
    End If
End Sub

Public Sub Pagto_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Trata_Troca_Aporte_TipoPagto(ByVal objCT As Object, ByVal lAporte As Long, ByVal iTipoPagto As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPagtoCond As ClassTRPAportePagtoCond
Dim objPagtoDireto As ClassTRPAportePagtoDireto

On Error GoTo Erro_Trata_Troca_Aporte_TipoPagto

    If lAporte <> objCT.gobjinfousu.lAporteAnt Or iTipoPagto <> objCT.gobjinfousu.iTipoPagtoAnt Then
    
        Pagto(objCT).Clear
        Set objCT.gobjinfousu.gobjPagtoCond = New ClassTRPAportePagtoCond
        Set objCT.gobjinfousu.gobjPagtoDireto = New ClassTRPAportePagtoDireto
        
        If iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
            
            iIndice = 0
            For Each objPagtoDireto In objCT.gobjinfousu.gobjAporte.colPagtoDireto
                iIndice = iIndice + 1
                'Se é sobre faturamento e ainda nãop foi faturado
                If objPagtoDireto.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then
                    Pagto(objCT).AddItem iIndice
                End If
            Next
        
        Else
        
            iIndice = 0
            For Each objPagtoCond In objCT.gobjinfousu.gobjAporte.colPagtoCondicionados
                iIndice = iIndice + 1
                'Se é sobre faturamento e ainda nãop foi faturado e já foi liberado
                If objPagtoCond.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT And objPagtoCond.iStatus = STATUS_TRP_OCR_LIBERADO Then
                    Pagto(objCT).AddItem iIndice
                End If
            Next
        
        End If
                
        objCT.gobjinfousu.lAporteAnt = lAporte
        objCT.gobjinfousu.iTipoPagtoAnt = iTipoPagto
    
    End If
    
    Trata_Troca_Aporte_TipoPagto = SUCESSO

    Exit Function

Erro_Trata_Troca_Aporte_TipoPagto:

    Trata_Troca_Aporte_TipoPagto = gErr

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192332)

    End Select

    Exit Function

End Function

Public Sub BotaoTrazerAporte_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objPagtoCond As ClassTRPAportePagtoCond
Dim objPagtoDireto As ClassTRPAportePagtoDireto
Dim iPagto As Integer
Dim objAporte As ClassTRPAportes
Dim sAporte As String
Dim dtData As Date
Dim dtDataReal As Date
Dim dValor As Double

On Error GoTo Erro_BotaoTrazerAporte_Click

    If Len(Aporte(objCT).Text) = 0 Then gError 192333
    If Len(TipoPagto(objCT).Text) = 0 Then gError 192334
    If Len(Pagto(objCT).Text) = 0 Then gError 192335
    
    sAporte = Aporte(objCT).Text
    
    objCT.gobjinfousu.giTrazendoAporte = MARCADO
    Call objCT.Limpa_Tela_NFFatPag
    objCT.gobjinfousu.giTrazendoAporte = DESMARCADO
    
    Aporte(gobjCT).PromptInclude = False
    Aporte(objCT).Text = sAporte
    Aporte(gobjCT).PromptInclude = True
    
    iPagto = StrParaInt(Len(Pagto(objCT).Text))
    
    Set objAporte = objCT.gobjinfousu.gobjAporte
    
    objCT.Fornecedor.Text = objAporte.lCliente
    Call objCT.Fornecedor_Validate(bSGECancelDummy)

    objCT.Filial.Text = objAporte.iFilialCliente
    Call objCT.Filial_Validate(bSGECancelDummy)
    
    objCT.Historico.Text = objAporte.sHistorico

    Call DateParaMasked(objCT.DataEmissao, gdtDataAtual)

    If objCT.gobjinfousu.iTipoPagtoAnt = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
        
        Set objPagtoDireto = objCT.gobjinfousu.gobjAporte.colPagtoDireto(iPagto)
    
        dValor = objPagtoDireto.dValor
        dtData = objPagtoDireto.dtVencimento
    
        Set objCT.gobjinfousu.gobjPagtoDireto = objPagtoDireto
        
    Else

        Set objPagtoCond = objCT.gobjinfousu.gobjAporte.colPagtoCondicionados(iPagto)
        
        dValor = objPagtoCond.dValor
        dtData = objPagtoCond.dtDataPagto
    
        Set objCT.gobjinfousu.gobjPagtoCond = objPagtoCond
    
    End If
    
    objCT.ValorTotal.Text = Format(dValor, "STANDARD")
    objCT.ValorProdutos.Text = Format(dValor, "STANDARD")
    
    
    'Calcula Data Vencimento Real
    lErro = CF("DataVencto_Real", dtData, dtDataReal)
    If lErro <> SUCESSO Then gError 192358

    objCT.GridParcelas.TextMatrix(1, objCT.iGrid_Parcela_Col) = "1"
    objCT.GridParcelas.TextMatrix(1, objCT.iGrid_Vencimento_Col) = Format(dtData, "dd/mm/yyyy")
    objCT.GridParcelas.TextMatrix(1, objCT.iGrid_VenctoReal_Col) = Format(dtDataReal, "dd/mm/yyyy")
    objCT.GridParcelas.TextMatrix(1, objCT.iGrid_ValorParcela_Col) = Format(dValor, "Standard")
    
    objCT.objGridParcelas.iLinhasExistentes = 1
    
    Exit Sub

Erro_BotaoTrazerAporte_Click:

    Select Case gErr
    
        Case 192333
             Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 192334
             Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_TIPOPAGTO_NAO_PREENCHIDO", gErr)
        
        Case 192335
             Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_PAGTO_NAO_PREENCHIDO", gErr)

        Case 192358

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192336)

    End Select

    Exit Sub

End Sub

Public Function NFFPAG_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Form_Load

    Set objEventoAporte = New AdmEvento
    
    Exit Function

Erro_NFFPAG_Form_Load:

    NFFPAG_Form_Load = gErr

    Select Case gErr
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192337)

    End Select

    Exit Function

End Function

Public Function NFFPAG_Form_UnLoad(ByVal objCT As Object, Cancel As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Form_UnLoad

    Set objEventoAporte = Nothing

    Exit Function

Erro_NFFPAG_Form_UnLoad:

    NFFPAG_Form_UnLoad = gErr

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192338)

    End Select

    Exit Function

End Function

Public Function NFFPAG_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.ActiveControl Is Aporte(objCT) Then
            Call objCT.LabelAporte_Click
        End If
    
    End If
    
    NFFPAG_UserControl_KeyDown = SUCESSO
    
End Function

Public Sub LabelAporte_Click(ByVal objCT As Object)

Dim colSelecao As New Collection
Dim objAporte As New ClassTRPAportes

    Set gobjCT = objCT
    
    Set objEventoAporte = New AdmEvento

    Call Chama_Tela("TRPAportesPagtosLista", colSelecao, objAporte, objEventoAporte, "StatusCod = 1 AND TipoDocDestino = 0 AND NumIntDocDestino = 0 AND FormaPagto = 1  AND TipoCod IN (1,2)")

End Sub

Private Sub objEventoAporte_evSelecao(obj1 As Object)

Dim objAporte As ClassTRPAportes
Dim lErro As Long

On Error GoTo Erro_objEventoAporte_evSelecao

    Set objAporte = obj1

    lErro = Traz_Aporte_Tela(gobjCT, objAporte)
    If lErro <> SUCESSO Then gError 192350

    gobjCT.Show
    
    Call BotaoTrazerAporte_Click(gobjCT)
    
    Exit Sub

Erro_objEventoAporte_evSelecao:

    Select Case gErr
    
        Case 192350

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192339)

    End Select

    Exit Sub

End Sub

Function NFFPAG_Traz_NFFatPag_Tela(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim objAporte As New ClassTRPAportes
Dim objPagtoCond As New ClassTRPAportePagtoCond
Dim objPagtoDireto As New ClassTRPAportePagtoDireto
Dim iTipoPagto As Integer
Dim lNumIntDocAporte As Long
Dim lNumIntDocPagto As Long

On Error GoTo Erro_NFFPAG_Traz_NFFatPag_Tela

    Set gobjCT = objCT

    iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO

    objPagtoDireto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
    objPagtoDireto.lNumIntDocDestino = objTituloPagar.lNumIntDoc

    lErro = CF("TRPAportePagtoDireto_Le_NumIntDocDestino", objPagtoDireto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192344
    
    lNumIntDocAporte = objPagtoDireto.lNumIntDocAporte
    lNumIntDocPagto = objPagtoDireto.lNumIntDoc
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then
    
        iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_COND
    
        objPagtoCond.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
        objPagtoCond.lNumIntDocDestino = objTituloPagar.lNumIntDoc
    
        lErro = CF("TRPAportePagtoCond_Le_NumIntDocDestino", objPagtoCond)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192345
    
        lNumIntDocAporte = objPagtoCond.lNumIntDocAporte
        lNumIntDocPagto = objPagtoCond.lNumIntDoc
    
    End If
       
    If lErro = SUCESSO Then
    
        objAporte.lNumIntDoc = lNumIntDocAporte
        
        'Lê o TRPAporte que está sendo Passado
        lErro = CF("TRPAportes_Le", objAporte)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192346
        
        objAporte.lNumIntDocPagto = lNumIntDocPagto
        objAporte.iTipoPagto = iTipoPagto
        
        lErro = Traz_Aporte_Tela(objCT, objAporte)
        If lErro <> SUCESSO Then gError 192351
    
    End If

    Exit Function

Erro_NFFPAG_Traz_NFFatPag_Tela:

    NFFPAG_Traz_NFFatPag_Tela = gErr

    Select Case gErr
    
        Case 192344 To 192346, 192351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192340)

    End Select

    Exit Function

End Function

Function NFFPAG_Move_Tela_Memoria(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long
    
Dim objAporte As New ClassTRPAportes
Dim iPagto As Integer
Dim iTipoPagto As Integer

On Error GoTo Erro_NFFPAG_Move_Tela_Memoria
    
    Set objAporte = objCT.gobjinfousu.gobjAporte
 
    If TipoPagto(objCT).ListIndex <> -1 Then 'And Len(Trim(Aporte(objCT).Text)) > 0 Then

        objAporte.iTipoPagto = TipoPagto(objCT).ItemData(TipoPagto(objCT).ListIndex)
        iPagto = StrParaInt(Pagto(objCT).Text)
        
        If objAporte.iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO And iPagto <> 0 Then
            objAporte.lNumIntDocPagto = objCT.gobjinfousu.gobjAporte.colPagtoDireto(iPagto).lNumIntDoc
        ElseIf iPagto <> 0 Then
            objAporte.lNumIntDocPagto = objCT.gobjinfousu.gobjAporte.colPagtoCondicionados(iPagto).lNumIntDoc
        End If
        
    End If
    
    Set objTituloPagar.objInfoUsu = objAporte
    
    NFFPAG_Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_NFFPAG_Move_Tela_Memoria:

    NFFPAG_Move_Tela_Memoria = gErr

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192357)

    End Select
    
    Exit Function
    
End Function

Function NFFPAG_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Limpa_Tela

    If objCT.gobjinfousu.giTrazendoAporte = DESMARCADO Then

        Set objCT.gobjinfousu.gobjAporte = New ClassTRPAportes
        Set objCT.gobjinfousu.gobjPagtoCond = New ClassTRPAportePagtoCond
        Set objCT.gobjinfousu.gobjPagtoDireto = New ClassTRPAportePagtoDireto
        
        objCT.gobjinfousu.lAporteAnt = 0
        objCT.gobjinfousu.iTipoPagtoAnt = 0
        objCT.gobjinfousu.iIndexPagtoAnt = 0
        
        TipoPagto(objCT).ListIndex = -1
        Pagto(objCT).Clear
        
    End If
    
    objCT.gobjinfousu.giTrazendoAporte = DESMARCADO
    
    NFFPAG_Limpa_Tela = SUCESSO

    Exit Function

Erro_NFFPAG_Limpa_Tela:

    objCT.gobjinfousu.giTrazendoAporte = DESMARCADO

    NFFPAG_Limpa_Tela = gErr

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192356)

    End Select
    
    Exit Function
    
End Function

Function NFFPAG_Gravar_Registro(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar, ByVal ColParcelaPagar As ColParcelaPagar) As Long

Dim lErro As Long
Dim objAporte As New ClassTRPAportes
Dim dValor As Double
Dim vbMsg As VbMsgBoxResult
Dim iPagto As Integer

On Error GoTo Erro_NFFPAG_Gravar_Registro

    Set objAporte = objTituloPagar.objInfoUsu

    If objAporte.lNumIntDocPagto > 0 Then
    
        'If objCT.gobjinfousu.iIndexPagtoAnt <> Pagto(objCT).ListIndex Then gError 192356
        
        If objAporte.lCliente <> objTituloPagar.lFornecedor Then gError 192357
        
        iPagto = StrParaInt(Len(Pagto(objCT).Text))
        
        If objAporte.iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
            dValor = objCT.gobjinfousu.gobjAporte.colPagtoDireto(iPagto).dValor
        Else
            dValor = objCT.gobjinfousu.gobjAporte.colPagtoCondicionados(iPagto).dValor
        End If
        
        If Abs(dValor - objTituloPagar.dValorTotal) > DELTA_VALORMONETARIO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "ERRO_VALOR_APORTE_DIF_VALOR_TITULO")
            If vbMsg = vbNo Then gError 192358
        
        End If
    
    End If
    
    NFFPAG_Gravar_Registro = SUCESSO

    Exit Function

Erro_NFFPAG_Gravar_Registro:

    NFFPAG_Gravar_Registro = gErr

    Select Case gErr
    
        Case 192357
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_APORTE_DIF_FORN_TITULO", gErr)

        Case 192358 'ERRO_CLIENTE_APORTE_DIF_FORN_TITULO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192353)

    End Select
    
End Function

Function Traz_Aporte_Tela(ByVal objCT As Object, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPagtoCond As ClassTRPAportePagtoCond
Dim objPagtoDireto As ClassTRPAportePagtoDireto

On Error GoTo Erro_Traz_Aporte_Tela

    Aporte(gobjCT).PromptInclude = False
    Aporte(gobjCT).Text = objAporte.lCodigo
    Aporte(gobjCT).PromptInclude = True
    
    Call Aporte_Validate(gobjCT, bSGECancelDummy)
    
    For iIndice = 0 To TipoPagto(gobjCT).ListCount - 1
        If TipoPagto(gobjCT).ItemData(iIndice) = objAporte.iTipoPagto Then
            TipoPagto(gobjCT).ListIndex = iIndice
            Exit For
        End If
    Next
    
    If objAporte.iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
        
        iIndice = -1
        For Each objPagtoDireto In objCT.gobjinfousu.gobjAporte.colPagtoDireto
        
            If objPagtoDireto.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then

                iIndice = iIndice + 1
                If objPagtoDireto.lNumIntDoc = objAporte.lNumIntDocPagto Then
                    Pagto(gobjCT).ListIndex = iIndice
                    Set objCT.gobjinfousu.gobjPagtoDireto = objPagtoDireto
                    Exit For
                End If
        
            End If
            
        Next
    
    Else
    
        iIndice = -1
        For Each objPagtoCond In objCT.gobjinfousu.gobjAporte.colPagtoCondicionados
            
            If objPagtoCond.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT And objPagtoCond.iStatus = STATUS_TRP_OCR_LIBERADO Then
            
                iIndice = iIndice + 1
                If objPagtoCond.lNumIntDoc = objAporte.lNumIntDocPagto Then
                    Pagto(gobjCT).ListIndex = iIndice
                    Set objCT.gobjinfousu.gobjPagtoCond = objPagtoCond
                    Exit For
                End If
        
            End If
            
        Next

    End If
    
    objCT.gobjinfousu.iIndexPagtoAnt = iIndice
    
    Traz_Aporte_Tela = SUCESSO

    Exit Function

Erro_Traz_Aporte_Tela:

    Traz_Aporte_Tela = gErr

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192343)

    End Select

    Exit Function

End Function
