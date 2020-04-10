VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTitPag_CMgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub BotaoCliente_Click(ByVal objCT As Object)
    Call objCT.gobjInfoUsu.BotaoCliente_Click(objCT, Me)
End Sub

Public Property Get Observacao(ByVal objCT As Object) As Object
     Set Observacao = objCT.objUserControl.Controls("Observacao")
End Property

Public Property Get Valor(ByVal objCT As Object) As Object
     Set Valor = objCT.objUserControl.Controls("Valor")
End Property

Public Property Get Cliente(ByVal objCT As Object) As Object
     Set Cliente = objCT.objUserControl.Controls("Cliente")
End Property

Public Property Get Cobrar(ByVal objCT As Object) As Object
     Set Cobrar = objCT.objUserControl.Controls("Cobrar")
End Property

Public Property Get Referencia(ByVal objCT As Object) As Object
     Set Referencia = objCT.objUserControl.Controls("Referencia")
End Property

Public Property Get Dossie(ByVal objCT As Object) As Object
     Set Dossie = objCT.objUserControl.Controls("Dossie")
End Property

Public Property Get ItemHistorico(ByVal objCT As Object) As Object
     Set ItemHistorico = objCT.objUserControl.Controls("ItemHistorico")
End Property

Public Property Get GridRateio(ByVal objCT As Object) As Object
     Set GridRateio = objCT.objUserControl.Controls("GridRateio")
End Property

Function TITPAGC_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_TITPAGC_Form_Load

    Call objCT.gobjInfoUsu.objEventoCliente_New
    Set objCT.gobjInfoUsu.objGridRateio = New AdmGrid

    lErro = Inicializa_Grid_Rateio(objCT, objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then gError 129018

    lErro = Carrega_Historico(objCT)
    If lErro <> SUCESSO Then gError 129050
    
    TITPAGC_Form_Load = SUCESSO
     
    Exit Function
    
Erro_TITPAGC_Form_Load:

    TITPAGC_Form_Load = gErr
     
    Select Case gErr
    
        Case 129018, 129050
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function TITPAGC_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    Set objCT.gobjInfoUsu.objGridRateio = Nothing
    Call objCT.gobjInfoUsu.objEventoCliente_Nothing
    
    TITPAGC_Form_Unload = SUCESSO
    
End Function

Function TITPAGC_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_TITPAGC_Limpa_Tela

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridRateio)

    TITPAGC_Limpa_Tela = SUCESSO
     
    Exit Function
    
Erro_TITPAGC_Limpa_Tela:

    TITPAGC_Limpa_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function TITPAGC_Traz_TituloPagar_Tela(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar) As Long
'Mostra na tela os dados da baixa

Dim lErro As Long
Dim colTituloPagRateio As New Collection
Dim objTituloPagRateio As ClassTituloPagRateio
Dim iLinha As Integer
Dim objCliente As ClassCliente
Dim iIndice As Integer
Dim iCodFilial As Integer

On Error GoTo Erro_TITPAGC_Traz_TituloPagar_Tela

    Observacao(objCT).Caption = objTituloPagar.sObservacao
    
    'Lê o Rateio vinculado ao Título
    lErro = CF("TituloPagarRateio_Le", objTituloPagar, colTituloPagRateio)
    If lErro <> SUCESSO Then gError 132269
    
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridRateio)
    
    iLinha = 0
 
    For Each objTituloPagRateio In colTituloPagRateio

        iLinha = iLinha + 1

        GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Valor_Col) = Format(objTituloPagRateio.dValor, "Standard")
              
        Cliente(objCT).Text = objTituloPagRateio.lCliente
        
        'Busca o Cliente no BD
        If Cliente(objCT).Text <> 0 Then
        
            Set objCliente = New ClassCliente
        
            lErro = TP_Cliente_Le3(Cliente(objCT), objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 132270

            GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Cliente_Col) = objCliente.sNomeReduzido
        End If
               
        For iIndice = 0 To ItemHistorico(objCT).ListCount - 1
            If ItemHistorico(objCT).ItemData(iIndice) = objTituloPagRateio.lHistorico Then
                GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Historico_Col) = ItemHistorico(objCT).List(iIndice)
                Exit For
            End If
        Next
        
        GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Dossie_Col) = RTrim(objTituloPagRateio.sDossie)
        GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Referencia_Col) = RTrim(objTituloPagRateio.sReferencia)
           
        GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Cobrar_Col) = objTituloPagRateio.iCobrar

    Next
    
    'Faz o número de linhas existentes do Grid ser igual ao número de Parcelas
    objCT.gobjInfoUsu.objGridRateio.iLinhasExistentes = iLinha
   
    Call Grid_Refresh_Checkbox(objCT.gobjInfoUsu.objGridRateio)
    
    TITPAGC_Traz_TituloPagar_Tela = SUCESSO

    Exit Function

Erro_TITPAGC_Traz_TituloPagar_Tela:

    TITPAGC_Traz_TituloPagar_Tela = gErr

    Select Case gErr

        Case 132269, 132270
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function TITPAGC_Saida_Celula(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_TITPAGC_Saida_Celula

    If objGridInt.objGrid Is GridRateio(objCT) Then

        Select Case objGridInt.objGrid.Col
    
            Case objCT.gobjInfoUsu.iGrid_Valor_Col
                lErro = Saida_Celula_ValorRateio(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132254
                
            Case objCT.gobjInfoUsu.iGrid_Cliente_Col
                lErro = Saida_Celula_Cliente(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132255
            
            Case objCT.gobjInfoUsu.iGrid_Historico_Col
                lErro = Saida_Celula_Historico(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132256
            
            Case objCT.gobjInfoUsu.iGrid_Referencia_Col
                lErro = Saida_Celula_Referencia(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132257
            
            Case objCT.gobjInfoUsu.iGrid_Dossie_Col
                lErro = Saida_Celula_Dossie(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132258
            
            Case objCT.gobjInfoUsu.iGrid_Cobrar_Col
                lErro = Saida_Celula_Cobrar(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132259
            
        End Select
        
    End If

    TITPAGC_Saida_Celula = SUCESSO

    Exit Function

Erro_TITPAGC_Saida_Celula:

    TITPAGC_Saida_Celula = gErr

    Select Case gErr

        Case 132254 To 132259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorRateio(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_ValorRateio

    Set objGridInt.objControle = Valor(objCT)

    'Verifica se valor está preenchido
    If Len(Trim(Valor(objCT).Text)) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(Valor(objCT).Text)
        If lErro <> SUCESSO Then gError 132261

        Valor(objCT).Text = Format(Valor(objCT).Text, "Standard")
        
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132262

    Saida_Celula_ValorRateio = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorRateio:

    Saida_Celula_ValorRateio = gErr

    Select Case gErr

        Case 132261, 132262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cliente(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cliente do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Saida_Celula_Cliente

    Set objGridInt.objControle = Cliente(objCT)
  
    'Verifica se Cliente está preenchido
    If Len(Trim(Cliente(objCT).Text)) > 0 Then
        
        'Busca o Cliente no BD
        lErro = TP_Cliente_Le3(Cliente(objCT), objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 132263

        Cliente(objCT).Text = objCliente.sNomeReduzido
                    
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132264

    Saida_Celula_Cliente = SUCESSO

    Exit Function

Erro_Saida_Celula_Cliente:

    Saida_Celula_Cliente = gErr

    Select Case gErr

        Case 132263, 132264
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobrar(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cobrar do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Cobrar

    Set objGridInt.objControle = Cobrar(objCT)

    'Verifica se valor está preenchido
    If Cobrar(objCT).Value <> 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132265

    Saida_Celula_Cobrar = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobrar:

    Saida_Celula_Cobrar = gErr

    Select Case gErr

        Case 132265
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Historico do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Historico

    Set objGridInt.objControle = ItemHistorico(objCT)

    'Verifica se Historico está preenchido
    If Len(Trim(ItemHistorico(objCT).Text)) > 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132266

    Saida_Celula_Historico = SUCESSO

    Exit Function

Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = gErr

    Select Case gErr

        Case 132266
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Dossie(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Dossie do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Dossie

    Set objGridInt.objControle = Dossie(objCT)

    'Verifica se Dossie está preenchido
    If Len(Trim(Dossie(objCT).Text)) > 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132267

    Saida_Celula_Dossie = SUCESSO

    Exit Function

Erro_Saida_Celula_Dossie:

    Saida_Celula_Dossie = gErr

    Select Case gErr

        Case 132267
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Referencia(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Referencia do grid que está deixando de ser a corrente
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Referencia

    Set objGridInt.objControle = Referencia(objCT)

    'Verifica se Referencia está preenchido
    If Len(Trim(Referencia(objCT).Text)) > 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132268

    Saida_Celula_Referencia = SUCESSO

    Exit Function

Erro_Saida_Celula_Referencia:

    Saida_Celula_Referencia = gErr

    Select Case gErr

        Case 132268
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub Valor_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Valor_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cliente_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cliente_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Referencia_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Referencia_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Dossie_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Dossie_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub ItemHistorico_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub ItemHistorico_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Valor_Validate(ByVal objCT As Object, Cancel As Boolean)
'Inserido por Wagner

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    Set objCT.gobjInfoUsu.objGridRateio.objControle = Valor(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Validate(ByVal objCT As Object, Cancel As Boolean) ' Inserido por Wagner

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate
   
    Set objCT.gobjInfoUsu.objGridRateio.objControle = Cliente(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub ItemHistorico_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ItemHistorico_Validate

    Set objCT.gobjInfoUsu.objGridRateio.objControle = ItemHistorico(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_ItemHistorico_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Dossie_Validate(ByVal objCT As Object, Cancel As Boolean) 'Inserido Por Wagner

Dim lErro As Long

On Error GoTo Erro_Dossie_Validate

    Set objCT.gobjInfoUsu.objGridRateio.objControle = Dossie(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_Dossie_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Referencia_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Referencia_Validate

    Set objCT.gobjInfoUsu.objGridRateio.objControle = Referencia(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_Referencia_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub GridRateio_Click(ByVal objCT As Object) 'Inserido por Wagner

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridRateio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)
    End If

End Sub

Public Sub GridRateio_GotFocus(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_EnterCell(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)

End Sub

Public Sub GridRateio_LeaveCell(ByVal objCT As Object) 'Inserido por Wagner

    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) 'Inserido por Wagner

    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_KeyPress(ByVal objCT As Object, KeyAscii As Integer) 'Inserido por Wagner

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridRateio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)
    End If

End Sub

Public Sub GridRateio_Validate(ByVal objCT As Object, Cancel As Boolean) 'Inserido por Wagner

    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_RowColChange(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_Scroll(ByVal objCT As Object) 'Inserido por Wagner

    Call Grid_Scroll(objCT.gobjInfoUsu.objGridRateio)

End Sub

Private Function Carrega_Historico(ByVal objCT As Object) As Long
'Carrega a combo de Historico
'Inserido por Wagner

Dim lErro As Long

On Error GoTo Erro_Carrega_Historico

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_DESCRICAORATEIO, ItemHistorico(objCT))
    If lErro <> SUCESSO Then gError 132273

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = gErr

    Select Case gErr
    
        Case 132273

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub Valor_Change(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Change(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrar_Click(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrar_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cobrar_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cobrar_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridRateio.objControle = Cobrar(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ItemHistorico_Change(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemHistorico_Click(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Dossie_Change(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Referencia_Change(ByVal objCT As Object) 'Inserido por Wagner

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Move_GridRateio_Memoria(ByVal objCT As Object, ColRateio As Collection, objTituloPagar As ClassTituloPagar) As Long
'Move para a memória os dados existentes no Grid
'Inserido por Wagner

Dim lErro As Long
Dim iIndice As Integer
Dim objTituloPagRateio As ClassTituloPagRateio
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_GridRateio_Memoria

    For iIndice = 1 To objCT.gobjInfoUsu.objGridRateio.iLinhasExistentes

        Set objTituloPagRateio = New ClassTituloPagRateio
        
        'Verifica se o cliente foi digitado
        If Len(Trim(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Cliente_Col))) > 0 Then
            objCliente.sNomeReduzido = GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Cliente_Col)

            'Lê o codigo através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 132274

            'Não achou o Cliente --> erro
            If lErro <> SUCESSO Then gError 132278

            'Guarda o código no objTituloReceber
            objTituloPagRateio.lCliente = objCliente.lCodigo
        End If

        'Preenche objTituloPagRateio com a linha do GridRateio
        objTituloPagRateio.dValor = StrParaDbl(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Valor_Col))
        objTituloPagRateio.iCobrar = StrParaInt(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Cobrar_Col))
        objTituloPagRateio.iSeq = iIndice
        objTituloPagRateio.lHistorico = Codigo_Extrai(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Historico_Col))
        objTituloPagRateio.lNumIntDocRec = 0
        objTituloPagRateio.sDossie = GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Dossie_Col)
        objTituloPagRateio.sReferencia = GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Referencia_Col)
        objTituloPagRateio.lNumIntDocPag = objTituloPagar.lNumIntDoc
       
        'Adiciona objTituloPagRateio à coleção colRateio
        ColRateio.Add objTituloPagRateio

    Next

    Move_GridRateio_Memoria = SUCESSO

    Exit Function

Erro_Move_GridRateio_Memoria:

    Move_GridRateio_Memoria = gErr

    Select Case gErr
    
        Case 132274
        
        Case 132275
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Cliente_Col))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro(objCT)
    If lErro <> SUCESSO Then gError 132276

    'Limpa a Tela
    Call objCT.Limpa_Tela_TituloPagar

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 132276

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro(ByVal objCT As Object) As Long
'Verifica se principais campos da tela foram preenchidos para gravação

Dim lErro As Long
Dim iIndice As Integer
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New ColParcelaPagar
Dim ColRateio As New Collection
Dim dSomaValorRateio As Double
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(objCT.Fornecedor.ClipText)) = 0 Then gError 132277
    If Len(Trim(objCT.Filial.Text)) = 0 Then gError 132278
    If Len(Trim(objCT.NumeroTitulo.Text)) = 0 Then gError 132279
    If Len(Trim(objCT.Tipo.Text)) = 0 Then gError 132280
    
    objFornecedor.sNomeReduzido = objCT.Fornecedor.Text

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 132284
    
    If lErro <> SUCESSO Then gError 132285
    
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(objCT.Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(objCT.Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(objCT.NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(objCT.DataEmissao.ClipText)) <> 0 Then
        objTituloPagar.dtDataEmissao = MaskedParaDate(objCT.DataEmissao)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 132281
    
    lErro = Move_GridRateio_Memoria(objCT, ColRateio, objTituloPagar)
    If lErro <> SUCESSO Then gError 132283
    
    dSomaValorRateio = 0
    For iIndice = 1 To objCT.gobjInfoUsu.objGridRateio.iLinhasExistentes
        dSomaValorRateio = dSomaValorRateio + StrParaDbl(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Valor_Col))
    Next
    
    If dSomaValorRateio <> objTituloPagar.dValorTotal And dSomaValorRateio <> 0 Then gError 132286
    
    Set objTituloPagar.objInfoUsu = ColRateio
    
    lErro = CF("TituloPagarRateio_Grava", objTituloPagar)
    If lErro <> SUCESSO Then gError 132282
    
    GL_objMDIForm.MousePointer = vbDefault
        
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 132277
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 132278
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 132279
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case 132280
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            
        Case 132281, 132282, 132283, 132284
                    
        Case 132285
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 132286
        'Inserido por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RATEIO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Rateio(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Referência")
    objGridInt.colColuna.Add ("Dossiê")
    objGridInt.colColuna.Add ("Cobrar")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Valor(objCT).Name)
    objGridInt.colCampo.Add (Cliente(objCT).Name)
    objGridInt.colCampo.Add (ItemHistorico(objCT).Name)
    objGridInt.colCampo.Add (Referencia(objCT).Name)
    objGridInt.colCampo.Add (Dossie(objCT).Name)
    objGridInt.colCampo.Add (Cobrar(objCT).Name)
        
    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_Rateio_Col = 0
    objCT.gobjInfoUsu.iGrid_Valor_Col = 1
    objCT.gobjInfoUsu.iGrid_Cliente_Col = 2
    objCT.gobjInfoUsu.iGrid_Historico_Col = 3
    objCT.gobjInfoUsu.iGrid_Referencia_Col = 4
    objCT.gobjInfoUsu.iGrid_Dossie_Col = 5
    objCT.gobjInfoUsu.iGrid_Cobrar_Col = 6
    
    'Grid do GridInterno
    objGridInt.objGrid = GridRateio(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 500

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6
    
    'Largura da primeira coluna
    GridRateio(objCT).ColWidth(0) = 300
    
    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    GridRateio(objCT).Width = 8000
        
    Inicializa_Grid_Rateio = SUCESSO

    Exit Function

End Function
