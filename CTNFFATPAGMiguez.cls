VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFATPAGMgz"
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

Public Property Get HistoricoGrid(ByVal objCT As Object) As Object
     Set HistoricoGrid = objCT.objUserControl.Controls("HistoricoGrid")
End Property

Public Property Get GridRateio(ByVal objCT As Object) As Object
     Set GridRateio = objCT.objUserControl.Controls("GridRateio")
End Property

Public Property Get Natureza(ByVal objCT As Object) As Object
     Set Natureza = objCT.objUserControl.Controls("Natureza")
End Property

Public Property Get LabelNatureza(ByVal objCT As Object) As Object
     Set LabelNatureza = objCT.objUserControl.Controls("LabelNatureza")
End Property

Public Property Get LabelNaturezaDesc(ByVal objCT As Object) As Object
     Set LabelNaturezaDesc = objCT.objUserControl.Controls("LabelNaturezaDesc")
End Property

Public Property Get Observacao(ByVal objCT As Object) As Object
     Set Observacao = objCT.objUserControl.Controls("Observacao")
End Property

Public Property Get Previsao(ByVal objCT As Object) As Object
     Set Previsao = objCT.objUserControl.Controls("Previsao")
End Property

Function NFFPAG_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Form_Load

    Call objCT.gobjInfoUsu.objEventoCliente_New
    Call objCT.gobjInfoUsu.objEventoNatureza_New
    
    Set objCT.gobjInfoUsu.objGridRateio = New AdmGrid

    lErro = Inicializa_Grid_Rateio(objCT, objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then gError 129018
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza(objCT)
    If lErro <> SUCESSO Then gError 32293

    lErro = Carrega_HistoricoGrid(objCT)
    If lErro <> SUCESSO Then gError 129050
    
    NFFPAG_Form_Load = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Form_Load:

    NFFPAG_Form_Load = gErr
     
    Select Case gErr
    
        Case 129018, 129050
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function NFFPAG_Gravar_Registro(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar, ByVal ColParcelaPagar As ColParcelaPagar) As Long

Dim lErro As Long
Dim dSomaValorRateio As Double
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTitPagPrevisao As New ClassTituloPagar
Dim objParcelaPag As ClassParcelaPagar

On Error GoTo Erro_NFFPAG_Gravar_Registro

    dSomaValorRateio = 0
    For iIndice = 1 To objCT.gobjInfoUsu.objGridRateio.iLinhasExistentes
        dSomaValorRateio = dSomaValorRateio + StrParaDbl(GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Valor_Col))
    Next
    
    If dSomaValorRateio <> objTituloPagar.dValorTotal And dSomaValorRateio <> 0 Then gError 129022

    If objTituloPagar.iStatus = STATUS_PREVISAO Then
        For Each objParcelaPag In ColParcelaPagar
            objParcelaPag.iStatus = STATUS_PREVISAO
        Next
    End If
                
    NFFPAG_Gravar_Registro = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Gravar_Registro:

    NFFPAG_Gravar_Registro = gErr
     
    Select Case gErr
        
        Case 129022
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RATEIO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function
    
End Function

Function NFFPAG_Gravar_Registro2(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar, ByVal ColParcelaPagar As ColParcelaPagar, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTitPagPrevisao As New ClassTituloPagar

On Error GoTo Erro_NFFPAG_Gravar_Registro2
        
    If Previsao(objCT).Value = vbUnchecked Then
    
        Set objTituloPagar.ColParcelaPagar = ColParcelaPagar
    
        'busca titulo com status previsao
        lErro = CF("TituloPag_BuscaPrevisao", objTituloPagar, objTitPagPrevisao)
        If lErro <> SUCESSO And lErro <> 106975 Then gError 32301
        
        If lErro = SUCESSO Then
        
            'pede confirmacao p/excluir
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_PREVISAO")
            If vbMsgRes = vbYes Then
            
                objTitPagPrevisao.iExclusao = 1
                
                'exclui previsao
                lErro = CF("NFFatPag_Exclui", objTitPagPrevisao, objContabil)
                If lErro <> SUCESSO Then gError 42785
            
            End If
            
        End If
    
    End If
        
    NFFPAG_Gravar_Registro2 = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Gravar_Registro2:

    NFFPAG_Gravar_Registro2 = gErr
     
    Select Case gErr

        Case 32301, 42785

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function
    
End Function

Function NFFPAG_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    Set objCT.gobjInfoUsu.objGridRateio = Nothing
    Call objCT.gobjInfoUsu.objEventoCliente_Nothing
    Call objCT.gobjInfoUsu.objEventoNatureza_Nothing
    
    NFFPAG_Form_Unload = SUCESSO
    
End Function

Function NFFPAG_Traz_NFFatPag_Tela(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim colTituloPagRateio As New Collection
Dim objTituloPagRateio As ClassTituloPagRateio
Dim iLinha As Integer
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim iIndice As Integer
Dim sNaturezaEnxuta As String

On Error GoTo Erro_NFFPAG_Traz_NFFatPag_Tela

    'Lê o Rateio vinculado ao Título
    lErro = CF("TituloPagarRateio_Le", objTituloPagar, colTituloPagRateio)
    If lErro <> SUCESSO Then gError 129047
    
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridRateio)
    
    iLinha = 0
 
    For Each objTituloPagRateio In colTituloPagRateio

        iLinha = iLinha + 1

        GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Valor_Col) = Format(objTituloPagRateio.dValor, "Standard")
              
        Cliente(objCT).Text = objTituloPagRateio.lCliente
        
        'Busca o Cliente no BD
        If Cliente(objCT).Text <> 0 Then
            lErro = TP_Cliente_Le3(Cliente(objCT), objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 129056

            GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Cliente_Col) = objCliente.sNomeReduzido
        End If
               
        For iIndice = 0 To HistoricoGrid(objCT).ListCount - 1
            If HistoricoGrid(objCT).ItemData(iIndice) = objTituloPagRateio.lHistorico Then
                GridRateio(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Historico_Col) = HistoricoGrid(objCT).List(iIndice)
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
    
    If Len(Trim(objTituloPagar.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objTituloPagar.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 32294
    
        Natureza(objCT).PromptInclude = False
        Natureza(objCT).Text = sNaturezaEnxuta
        Natureza(objCT).PromptInclude = True
        
    Else
    
        Natureza(objCT).PromptInclude = False
        Natureza(objCT).Text = ""
        Natureza(objCT).PromptInclude = True
        
    End If
    
    Call Natureza_Validate(objCT, bSGECancelDummy)
    
    Observacao(objCT).Text = objTituloPagar.sObservacao
    
    If objTituloPagar.iStatus = STATUS_PREVISAO Then
        Previsao(objCT).Value = vbChecked
    Else
        Previsao(objCT).Value = vbUnchecked
    End If
    
    NFFPAG_Traz_NFFatPag_Tela = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Traz_NFFatPag_Tela:

    NFFPAG_Traz_NFFatPag_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function NFFPAG_Saida_Celula(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Saida_Celula

    If objGridInt.objGrid Is GridRateio(objCT) Then

        Select Case objGridInt.objGrid.Col
    
            Case objCT.gobjInfoUsu.iGrid_Valor_Col
                lErro = Saida_Celula_ValorRateio(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129024
                
            Case objCT.gobjInfoUsu.iGrid_Cliente_Col
                lErro = Saida_Celula_Cliente(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129025
            
            Case objCT.gobjInfoUsu.iGrid_Historico_Col
                lErro = Saida_Celula_HistoricoGrid(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129026
            
            Case objCT.gobjInfoUsu.iGrid_Referencia_Col
                lErro = Saida_Celula_Referencia(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129027
            
            Case objCT.gobjInfoUsu.iGrid_Dossie_Col
                lErro = Saida_Celula_Dossie(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129028
            
            Case objCT.gobjInfoUsu.iGrid_Cobrar_Col
                lErro = Saida_Celula_Cobrar(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 129029
            
        End Select
        
    End If

    NFFPAG_Saida_Celula = SUCESSO

    Exit Function

Erro_NFFPAG_Saida_Celula:

    NFFPAG_Saida_Celula = gErr

    Select Case gErr

        Case 129024 To 129029

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function NFFPAG_Move_Tela_Memoria(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim ColRateio As New Collection
Dim sNaturezaFormatada  As String
Dim iNaturezaPreenchida As Integer

On Error GoTo Erro_NFFPAG_Move_Tela_Memoria

    lErro = Move_GridRateio_Memoria(objCT, ColRateio, objTituloPagar)
    If lErro <> SUCESSO Then gError 132314
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza(objCT).Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 32295
    
    objTituloPagar.sNatureza = sNaturezaFormatada
    
    objTituloPagar.sObservacao = Observacao(objCT).Text
    If Previsao(objCT).Value = vbChecked Then
        objTituloPagar.iStatus = STATUS_PREVISAO
    Else
        objTituloPagar.iStatus = STATUS_LANCADO
    End If
        
    NFFPAG_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Move_Tela_Memoria:

    NFFPAG_Move_Tela_Memoria = gErr
     
    Select Case gErr
    
        Case 132314
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function NFFPAG_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long

    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.ActiveControl Is Cliente(objCT) Then
            Call BotaoCliente_Click(objCT)
        End If
    
    End If
    
    NFFPAG_UserControl_KeyDown = SUCESSO
    
End Function

Function NFFPAG_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Limpa_Tela

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridRateio)
    
    Natureza(objCT).PromptInclude = False
    Natureza(objCT).Text = ""
    Natureza(objCT).PromptInclude = True
    
    LabelNaturezaDesc(objCT).Caption = ""
    
    Previsao(objCT).Value = vbUnchecked
    
    NFFPAG_Limpa_Tela = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Limpa_Tela:

    NFFPAG_Limpa_Tela = gErr
     
    Select Case gErr
          
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
        If lErro <> SUCESSO Then gError 129019

        Valor(objCT).Text = Format(Valor(objCT).Text, "Standard")
        
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129032

    Saida_Celula_ValorRateio = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorRateio:

    Saida_Celula_ValorRateio = gErr

    Select Case gErr

        Case 129019, 129032
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cliente(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cliente do grid que está deixando de ser a corrente

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
        If lErro <> SUCESSO Then gError 129053

        Cliente(objCT).Text = objCliente.sNomeReduzido
                    
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129033

    Saida_Celula_Cliente = SUCESSO

    Exit Function

Erro_Saida_Celula_Cliente:

    Saida_Celula_Cliente = gErr

    Select Case gErr

        Case 129033
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129053
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobrar(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cobrar do grid que está deixando de ser a corrente

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
    If lErro <> SUCESSO Then gError 129034

    Saida_Celula_Cobrar = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobrar:

    Saida_Celula_Cobrar = gErr

    Select Case gErr

        Case 129034
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HistoricoGrid(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula HistoricoGrid do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_HistoricoGrid

    Set objGridInt.objControle = HistoricoGrid(objCT)

    'Verifica se HistoricoGrid está preenchido
    If Len(Trim(HistoricoGrid(objCT).Text)) > 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridRateio(objCT).Row - GridRateio(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129035

    Saida_Celula_HistoricoGrid = SUCESSO

    Exit Function

Erro_Saida_Celula_HistoricoGrid:

    Saida_Celula_HistoricoGrid = gErr

    Select Case gErr

        Case 129035
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Dossie(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Dossie do grid que está deixando de ser a corrente

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
    If lErro <> SUCESSO Then gError 129036

    Saida_Celula_Dossie = SUCESSO

    Exit Function

Erro_Saida_Celula_Dossie:

    Saida_Celula_Dossie = gErr

    Select Case gErr

        Case 129036
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Referencia(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da celula Referencia do grid que está deixando de ser a corrente

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
    If lErro <> SUCESSO Then gError 129037

    Saida_Celula_Referencia = SUCESSO

    Exit Function

Erro_Saida_Celula_Referencia:

    Saida_Celula_Referencia = gErr

    Select Case gErr

        Case 129037
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub Valor_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Valor_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cliente_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Cliente_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Referencia_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Referencia_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Dossie_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Dossie_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub HistoricoGrid_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub HistoricoGrid_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub Valor_Validate(ByVal objCT As Object, Cancel As Boolean)

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

Public Sub Cliente_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer, bCancel As Boolean
Dim colCodigoNome As New AdmColCodigoNome

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

Public Sub HistoricoGrid_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HistoricoGrid_Validate

    Set objCT.gobjInfoUsu.objGridRateio.objControle = HistoricoGrid(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub

Erro_HistoricoGrid_Validate:

    Cancel = True

    Select Case gErr


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Dossie_Validate(ByVal objCT As Object, Cancel As Boolean)

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

Public Sub GridRateio_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridRateio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)
    End If

End Sub

Public Sub GridRateio_GotFocus(ByVal objCT As Object)

    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_EnterCell(ByVal objCT As Object)

    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)

End Sub

Public Sub GridRateio_LeaveCell(ByVal objCT As Object)

    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridRateio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridRateio, objCT.iAlterado)
    End If

End Sub

Public Sub GridRateio_Validate(ByVal objCT As Object, Cancel As Boolean)

    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_RowColChange(ByVal objCT As Object)

    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridRateio)

End Sub

Public Sub GridRateio_Scroll(ByVal objCT As Object)

    Call Grid_Scroll(objCT.gobjInfoUsu.objGridRateio)

End Sub

Private Function Carrega_HistoricoGrid(ByVal objCT As Object) As Long
'Carrega a combo de HistoricoGrid

Dim lErro As Long

On Error GoTo Erro_Carrega_HistoricoGrid

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_DESCRICAORATEIO, HistoricoGrid(objCT))
    If lErro <> SUCESSO Then gError 129056

    Carrega_HistoricoGrid = SUCESSO

    Exit Function

Erro_Carrega_HistoricoGrid:

    Carrega_HistoricoGrid = gErr

    Select Case gErr
    
        Case 129056

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
    objGridInt.colCampo.Add (HistoricoGrid(objCT).Name)
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
    objGridInt.iLinhasVisiveis = 10
    
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

Public Sub Valor_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrar_Click(ByVal objCT As Object)

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

Public Sub HistoricoGrid_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub HistoricoGrid_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Dossie_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Referencia_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Move_GridRateio_Memoria(ByVal objCT As Object, ByVal ColRateio As Collection, ByVal objTituloPagar As ClassTituloPagar) As Long
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
            If lErro <> SUCESSO And lErro <> 12348 Then gError 129051

            'Não achou o Cliente --> erro
            If lErro <> SUCESSO Then gError 129052

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
    
    Set objTituloPagar.objInfoUsu = ColRateio

    Move_GridRateio_Memoria = SUCESSO

    Exit Function

Erro_Move_GridRateio_Memoria:

    Move_GridRateio_Memoria = gErr

    Select Case gErr
    
        Case 129051
        
        Case 129052
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, GridRateio(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Cliente_Col))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Sub Natureza_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Sub Natureza_Validate(ByVal objCT As Object, Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza(objCT).ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza(objCT).Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_PAGAMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc(objCT).Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc(objCT).Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub

Private Function Inicializa_Mascara_Natureza(ByVal objCT As Object) As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza(objCT).Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

