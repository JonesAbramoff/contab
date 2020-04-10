VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFormataExt2"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Funcao Copiada
Public Function Ordena_Colecao(colEntrada As Collection, colSaida As Collection, colCampos As Collection) As Long

Dim iIndice As Long
Dim lPontoSuperior As Long
Dim lPontoInferior As Long
Dim lPontoMedio As Long
Dim iResto As Integer
Dim iResultado1 As Integer
Dim iResultado2 As Integer
Dim iResultado As Integer
Dim lErro As Long

On Error GoTo Erro_Ordena_Colecao

    'Para cada Item da Colecao de Entrada
    For iIndice = 1 To colEntrada.Count

        'Inicializa as Variaveis
        lPontoSuperior = colSaida.Count
        lPontoInferior = 1
        lPontoMedio = 1

        'Se a Colecao está Vazia
        If colSaida.Count = 0 Then

            'Adiciona em qualquer posicao
            Call colSaida.Add(colEntrada(iIndice))

        'Se a colecao só tem um Item
        ElseIf colSaida.Count = 1 Then

            'Compara para saber a posicao
            iResultado = Comparacao(colEntrada(iIndice), colSaida(1), colCampos)

            'se está acima do Item da colecao
            If iResultado = MENOR Then
                'Adiciona antes
                Call colSaida.Add(colEntrada(iIndice), , 1)
            Else
                'adiciona depois do Item
                Call colSaida.Add(colEntrada(iIndice), , , 1)
            End If

        'Se na Colecao só tem 2 Itens
        ElseIf colSaida.Count = 2 Then

            'Verifica onde está e coloca posicionado
            iResultado1 = Comparacao(colEntrada(iIndice), colSaida(1), colCampos)
            iResultado2 = Comparacao(colEntrada(iIndice), colSaida(colSaida.Count), colCampos)

            If iResultado1 = MENOR Then
                Call colSaida.Add(colEntrada(iIndice), , 1)
            ElseIf iResultado1 = MAIOR Then
                If iResultado2 = MENOR Then
                    Call colSaida.Add(colEntrada(iIndice), , colSaida.Count)
                Else
                    Call colSaida.Add(colEntrada(iIndice), , , colSaida.Count)
                End If
            ElseIf iResultado1 = IGUAL Then
                Call colSaida.Add(colEntrada(iIndice), , 1)
            End If

        Else

            'Compara se está acima do Limite Inferior da colecao de Saida
            iResultado1 = Comparacao(colEntrada(iIndice), colSaida(1), colCampos)

            'Compara se está abaixo do Limite Superior da colecao de Saída
            iResultado2 = Comparacao(colEntrada(iIndice), colSaida(colSaida.Count), colCampos)

            'Se está dentro dos Limites
            If iResultado1 = MAIOR And iResultado2 = MENOR Then

                'Faz Procura Logarítima
                Do While ((lPontoSuperior - lPontoInferior) > 1)

                    'Calcula o Ponto Médio
                    iResto = Resto((lPontoSuperior + lPontoInferior), 2)

                    If iResto = 0 Then
                        lPontoMedio = (lPontoSuperior + lPontoInferior) / 2
                    Else
                        lPontoMedio = (lPontoSuperior + lPontoInferior + 1) / 2
                    End If

                    'Compara com o ponto médio
                    iResultado = Comparacao(colEntrada(iIndice), colSaida(lPontoMedio), colCampos)

                    'Se fo igual
                    If iResultado = IGUAL Then
                        'Sai e adiciona na colecao
                        Exit Do
                    ElseIf iResultado = MAIOR Then
                        'Troca o valor do Ponto inferior para ser o Ponto Médio
                        lPontoInferior = lPontoMedio
                    ElseIf iResultado = MENOR Then
                        'Troca o valor do Ponto Superior para ser o Ponto Médio
                        lPontoSuperior = lPontoMedio
                    End If
                Loop
                
                If iResultado = MENOR Then
                    'Adiciona antes
                    Call colSaida.Add(colEntrada(iIndice), , lPontoMedio)
                Else
                    'Adiciona antes
                    Call colSaida.Add(colEntrada(iIndice), , , lPontoMedio)
                End If
                
            Else
                
                'para valores acima do limite inferior
                If iResultado1 <> MAIOR Then
                    'Adiciona antes do Primeiro na coleção de Saída
                    Call colSaida.Add(colEntrada(iIndice), , 1)
                'para valores abaixo do limite Superior
                ElseIf iResultado2 <> MENOR Then
                    'Adiciona depois do último na colção de Saída
                    Call colSaida.Add(colEntrada(iIndice), , , colSaida.Count)
                End If

            End If

        End If
    Next

    Exit Function

Erro_Ordena_Colecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142373)

    End Select

    Exit Function

End Function

Private Function Comparacao(objEntrada As Object, objMedio As Object, colCampos As Collection) As Integer
'Compara os campos do objEntrada e do objMedio e retorna o Resultado em colCampos

Dim iIndice As Integer
Dim sProcName As String
Dim iIndiceCampo As Integer
Dim iPos As Integer

    Comparacao = IGUAL

    For iIndice = 1 To colCampos.Count
    
        iPos = InStr(1, colCampos(iIndice), "(")
    
        If iPos = 0 Then
            sProcName = colCampos(iIndice)
            iIndiceCampo = 0
        Else
            sProcName = left(colCampos(iIndice), iPos - 1)
            iIndiceCampo = CInt(Mid(colCampos(iIndice), iPos + 1, Len(colCampos(iIndice)) - 1 - iPos))
        End If
        
        If iIndiceCampo = 0 Then

            If CallByName(objEntrada, sProcName, VbGet) > CallByName(objMedio, sProcName, VbGet) Then
                Comparacao = MAIOR
                Exit For
            ElseIf CallByName(objEntrada, sProcName, VbGet) < CallByName(objMedio, sProcName, VbGet) Then
                Comparacao = MENOR
                Exit For
            End If
            
        Else
        
            If CallByName(objEntrada, sProcName, VbGet, iIndiceCampo) > CallByName(objMedio, sProcName, VbGet, iIndiceCampo) Then
                Comparacao = MAIOR
                Exit For
            ElseIf CallByName(objEntrada, sProcName, VbGet, iIndiceCampo) < CallByName(objMedio, sProcName, VbGet, iIndiceCampo) Then
                Comparacao = MENOR
                Exit For
            End If
        
        End If

    Next

    Exit Function

End Function

Public Function Ordena_Grid(ByVal objGridInt As AdmGrid, ByVal colColunas As Collection, ByVal iOrdemDesc As Integer, ByVal colcolColecoesOrd As Collection) As Long
'Ordena um Grid em uma coluna

Dim lErro As Long
Dim iLinha As Integer
Dim iColuna As Integer
Dim iIndice As Integer
Dim objOrdenaColecao As ClassOrdenaColecao
Dim colOrdenaColecao As New Collection
Dim colSaida As New Collection
Dim colSaidaInv As New Collection
Dim colCampos As New Collection
Dim iPos As Integer
Dim sValorCol As String
Dim colTipoColuna As New Collection
Dim vTipo As Variant
Dim vTipoVal As Variant
Dim colColecoesOrd As New Collection
Dim colColecoesOrdAux As New Collection
Dim vColuna As Variant

Const COLUNA_TIPO_TEXTO = 1
Const COLUNA_TIPO_TEXTO_SEP = 2
Const COLUNA_TIPO_NUMERO = 3
Const COLUNA_TIPO_NUMERO_SEP = 4
Const COLUNA_TIPO_DATA = 5

On Error GoTo Erro_Ordena_Grid

    'Obtém os tipos de campos do grid
    For iColuna = 1 To objGridInt.objGrid.Cols - 1
        vTipoVal = 0
        For iLinha = 1 To objGridInt.iLinhasExistentes
        
            sValorCol = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
            
            iPos = InStr(1, sValorCol, SEPARADOR)
            If iPos <> 0 Then
                sValorCol = SCodigo_Extrai(sValorCol)
            End If
            If IsNumeric(sValorCol) Then
                If iPos = 0 Then
                    vTipo = COLUNA_TIPO_NUMERO
                Else
                    vTipo = COLUNA_TIPO_NUMERO_SEP
                End If
            Else
                If IsDate(sValorCol) Then
                    vTipo = COLUNA_TIPO_DATA
                Else
                    If iPos = 0 Then
                        vTipo = COLUNA_TIPO_TEXTO
                    Else
                        vTipo = COLUNA_TIPO_TEXTO_SEP
                    End If
                End If
            End If
            
            If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iColuna))) <> 0 Then
            
                Select Case vTipo
                
                    Case COLUNA_TIPO_TEXTO
                        vTipoVal = COLUNA_TIPO_TEXTO
                
                    Case COLUNA_TIPO_TEXTO_SEP
                    
                        If vTipoVal = 0 Or vTipoVal = COLUNA_TIPO_NUMERO_SEP Or vTipoVal = COLUNA_TIPO_TEXTO_SEP Then
                            vTipoVal = COLUNA_TIPO_TEXTO_SEP
                        Else
                            vTipoVal = COLUNA_TIPO_TEXTO
                        End If
                
                    Case COLUNA_TIPO_NUMERO
                
                        If vTipoVal = 0 Or vTipoVal = COLUNA_TIPO_NUMERO Then
                            vTipoVal = COLUNA_TIPO_NUMERO
                        Else
                            vTipoVal = COLUNA_TIPO_TEXTO
                        End If
                        
                    Case COLUNA_TIPO_NUMERO_SEP
                
                        If vTipoVal = 0 Or vTipoVal = COLUNA_TIPO_NUMERO_SEP Then
                            vTipoVal = COLUNA_TIPO_NUMERO_SEP
                        Else
                            vTipoVal = COLUNA_TIPO_TEXTO
                        End If
                        
                    Case COLUNA_TIPO_DATA
                
                        If vTipoVal = 0 Or vTipoVal = COLUNA_TIPO_DATA Then
                            vTipoVal = COLUNA_TIPO_DATA
                        Else
                            vTipoVal = COLUNA_TIPO_TEXTO
                        End If
                        
                End Select
            
            End If
            
        Next
        If vTipoVal = 0 Then vTipoVal = COLUNA_TIPO_TEXTO
        colTipoColuna.Add vTipoVal
    Next

    'Joga os itens do Grid para uma coleção generica
    For iLinha = 1 To objGridInt.iLinhasExistentes
    
        Set objOrdenaColecao = New ClassOrdenaColecao
        
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
        
            vTipo = colTipoColuna(iColuna)
        
            sValorCol = objGridInt.objGrid.TextMatrix(iLinha, iColuna)
            objOrdenaColecao.sColuna(iColuna) = sValorCol
            
            Select Case vTipo
            
                Case COLUNA_TIPO_TEXTO
                    objOrdenaColecao.vColuna(iColuna) = sValorCol
            
                Case COLUNA_TIPO_TEXTO_SEP
                    sValorCol = SCodigo_Extrai(sValorCol)
                    objOrdenaColecao.vColuna(iColuna) = sValorCol
            
                Case COLUNA_TIPO_NUMERO
                    If Len(Trim(sValorCol)) <> 0 Then
                        objOrdenaColecao.vColuna(iColuna) = CDbl(sValorCol)
                    Else
                        objOrdenaColecao.vColuna(iColuna) = 0
                    End If
                    
                Case COLUNA_TIPO_NUMERO_SEP
                    sValorCol = SCodigo_Extrai(sValorCol)
                    If Len(Trim(sValorCol)) <> 0 Then
                        objOrdenaColecao.vColuna(iColuna) = CDbl(sValorCol)
                    Else
                        objOrdenaColecao.vColuna(iColuna) = 0
                    End If
                    
                Case COLUNA_TIPO_DATA
                    If Len(Trim(sValorCol)) <> 0 Then
                        objOrdenaColecao.vColuna(iColuna) = CDate(sValorCol)
                    Else
                        objOrdenaColecao.vColuna(iColuna) = DATA_NULA
                    End If
                    
            End Select
            
        Next
        
        objOrdenaColecao.iIndice = iLinha
        colOrdenaColecao.Add objOrdenaColecao
    Next
    
    For Each vColuna In colColunas
        colCampos.Add "vColuna(" & vColuna & ")"
    Next

    'Ordena a coleção
    lErro = Ordena_Colecao(colOrdenaColecao, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 182170
    
    'Inverte a ordenação se for o caso
    If iOrdemDesc <> ORDEM_CRESCENTE Then
        For iIndice = colSaida.Count To 1 Step -1
            colSaidaInv.Add colSaida.Item(iIndice)
        Next
        Set colSaida = colSaidaInv
    End If
       
    'Reordena as coleções passadas
    If Not (colcolColecoesOrd Is Nothing) Then
        For Each colColecoesOrd In colcolColecoesOrd
        
            Set colColecoesOrdAux = New Collection
        
            For Each objOrdenaColecao In colSaida
                colColecoesOrdAux.Add colColecoesOrd.Item(objOrdenaColecao.iIndice)
            Next
            
            For iIndice = colColecoesOrd.Count To 1 Step -1
                colColecoesOrd.Remove iIndice
            Next
        
            For iIndice = 1 To colColecoesOrdAux.Count
                colColecoesOrd.Add colColecoesOrdAux.Item(iIndice)
            Next
        
        Next
    End If
    
    Call Grid_Limpa(objGridInt)

    'Transfere da coleção ordenada para o Grid
    iLinha = 0
    For Each objOrdenaColecao In colSaida
        iLinha = iLinha + 1
        For iColuna = 1 To objGridInt.objGrid.Cols - 1
            objGridInt.objGrid.TextMatrix(iLinha, iColuna) = objOrdenaColecao.sColuna(iColuna)
        Next
    Next
    
    Call Grid_Refresh_Checkbox(objGridInt)
    
    objGridInt.iLinhasExistentes = iLinha
    
    Ordena_Grid = SUCESSO

    Exit Function

Erro_Ordena_Grid:

    Ordena_Grid = gErr

    Select Case gErr
    
        Case 182170

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182171)

    End Select

    Exit Function
    
End Function

Public Function Ordenacao_ClickGrid(objGridInt As AdmGrid, Optional objComboOrdenacao As Object, Optional colcolColecoes As Collection) As Long
'Ordena o Grid pelo Click do Mouse na coluna 0

Dim lErro As Long
Dim iMouseRow As Integer
Dim sColuna As String
Dim iColuna As Integer
Dim colColunas As New Collection

On Error GoTo Erro_Ordenacao_ClickGrid

    iMouseRow = objGridInt.objGrid.MouseRow

    If iMouseRow = 0 Then
    
        If Len(Trim(objGridInt.objGrid.TextMatrix(0, objGridInt.objGrid.Col))) > 0 Then
        
            For iColuna = 1 To objGridInt.objGrid.Cols - 1
           
                If iColuna <> objGridInt.objGrid.Col Then
                    sColuna = objGridInt.objGrid.TextMatrix(0, iColuna)
                    
                    If Len(Trim(sColuna)) > 1 Then
                    
                        If right(sColuna, 1) = SETA_PARA_CIMA Or right(sColuna, 1) = SETA_PARA_BAIXO Then
                        
                            objGridInt.objGrid.TextMatrix(0, iColuna) = left(sColuna, Len(sColuna) - 2)
                        
                        End If
                    
                    End If
                    
                End If
            
            Next
            
            colColunas.Add objGridInt.objGrid.Col
            sColuna = objGridInt.objGrid.TextMatrix(0, objGridInt.objGrid.Col)
    
            If right(sColuna, 1) = SETA_PARA_CIMA Then
                objGridInt.objGrid.TextMatrix(0, objGridInt.objGrid.Col) = left(sColuna, Len(sColuna) - 1) & SETA_PARA_BAIXO
                
                If Not (objComboOrdenacao Is Nothing) Then
                    Call ListOuCombo_Seleciona_ItemData(objComboOrdenacao, -1 * objGridInt.objGrid.Col)
                Else
                    Call Ordena_Grid(objGridInt, colColunas, ORDEM_DECRESCENTE, colcolColecoes)
                End If
            
            ElseIf right(sColuna, 1) = SETA_PARA_BAIXO Then
                objGridInt.objGrid.TextMatrix(0, objGridInt.objGrid.Col) = left(sColuna, Len(sColuna) - 1) & SETA_PARA_CIMA
            
                If Not (objComboOrdenacao Is Nothing) Then
                    Call ListOuCombo_Seleciona_ItemData(objComboOrdenacao, objGridInt.objGrid.Col)
                Else
                    Call Ordena_Grid(objGridInt, colColunas, ORDEM_CRESCENTE, colcolColecoes)
                End If
                
            Else
                objGridInt.objGrid.TextMatrix(0, objGridInt.objGrid.Col) = sColuna & " " & SETA_PARA_CIMA
                
                If Not (objComboOrdenacao Is Nothing) Then
                    Call ListOuCombo_Seleciona_ItemData(objComboOrdenacao, objGridInt.objGrid.Col)
                Else
                    Call Ordena_Grid(objGridInt, colColunas, ORDEM_CRESCENTE, colcolColecoes)
                End If
            
            End If
            
        End If
    
    End If
    
    Ordenacao_ClickGrid = SUCESSO

    Exit Function

Erro_Ordenacao_ClickGrid:

    Ordenacao_ClickGrid = gErr

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182260)

    End Select

    Exit Function
    
End Function

Public Function Ordenacao_Limpa(objGridInt As AdmGrid, Optional objComboOrdenacao As Object) As Long
'Ordena o Grid pelo Click do Mouse na coluna 0

Dim lErro As Long
Dim sColuna As String
Dim iColuna As Integer

On Error GoTo Erro_Ordenacao_Limpa

    For iColuna = 1 To objGridInt.objGrid.Cols - 1
             
        sColuna = objGridInt.objGrid.TextMatrix(0, iColuna)
        
        If Len(Trim(sColuna)) > 1 Then
        
            If right(sColuna, 1) = SETA_PARA_CIMA Or right(sColuna, 1) = SETA_PARA_BAIXO Then
            
                objGridInt.objGrid.TextMatrix(0, iColuna) = left(sColuna, Len(sColuna) - 2)
            
            End If
        
        End If
     
    Next
    
    If Not (objComboOrdenacao Is Nothing) Then
        objComboOrdenacao.ListIndex = -1
    End If
    
    Ordenacao_Limpa = SUCESSO

    Exit Function

Erro_Ordenacao_Limpa:

    Ordenacao_Limpa = gErr

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182260)

    End Select

    Exit Function
    
End Function

Public Function Ordenacao_Atualiza(objGridInt As AdmGrid, objComboOrdenacao As Object, Optional ByVal colcolColecoes As Collection) As Long
'Ordena o Grid pelo Click do Mouse na coluna 0

Dim lErro As Long
Dim colColunas As New Collection
Dim iOrdem As Integer
Dim iColuna As Integer
Dim sColuna As String

On Error GoTo Erro_Ordenacao_Atualiza

    For iColuna = 1 To objGridInt.objGrid.Cols - 1
             
        sColuna = objGridInt.objGrid.TextMatrix(0, iColuna)
        
        If Len(Trim(sColuna)) > 1 Then
        
            If right(sColuna, 1) = SETA_PARA_CIMA Or right(sColuna, 1) = SETA_PARA_BAIXO Then
            
                objGridInt.objGrid.TextMatrix(0, iColuna) = left(sColuna, Len(sColuna) - 2)
            
            End If
        
        End If
     
    Next
    
    If objComboOrdenacao.ListIndex <> -1 Then
        
        'Se for crescente
        If objComboOrdenacao.ItemData(objComboOrdenacao.ListIndex) > 0 Then
    
            iColuna = objComboOrdenacao.ItemData(objComboOrdenacao.ListIndex)
    
            colColunas.Add objComboOrdenacao.ItemData(objComboOrdenacao.ListIndex)
            Call Ordena_Grid(objGridInt, colColunas, ORDEM_CRESCENTE, colcolColecoes)
            
            objGridInt.objGrid.TextMatrix(0, iColuna) = objGridInt.objGrid.TextMatrix(0, iColuna) & " " & SETA_PARA_CIMA
            
        Else
            iColuna = -objComboOrdenacao.ItemData(objComboOrdenacao.ListIndex)
        
            colColunas.Add iColuna
            Call Ordena_Grid(objGridInt, colColunas, ORDEM_DECRESCENTE, colcolColecoes)
        
            objGridInt.objGrid.TextMatrix(0, iColuna) = objGridInt.objGrid.TextMatrix(0, iColuna) & " " & SETA_PARA_BAIXO
        
        End If
        
    End If
    
    Ordenacao_Atualiza = SUCESSO

    Exit Function

Erro_Ordenacao_Atualiza:

    Ordenacao_Atualiza = gErr

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182260)

    End Select

    Exit Function
    
End Function

Public Function Ordenacao_Preeenche(objGridInt As AdmGrid, objComboOrdenacao As Object) As Long
'Ordena o Grid pelo Click do Mouse na coluna 0

Dim lErro As Long
Dim iIndice As Integer
Dim vColuna As Variant

On Error GoTo Erro_Ordenacao_Preeenche

    objComboOrdenacao.Clear

    'Atualiza a Parte de Ordenação
    iIndice = -1
    For Each vColuna In objGridInt.colColuna
        iIndice = iIndice + 1
        If iIndice <> 0 Then
            objComboOrdenacao.AddItem vColuna & SEPARADOR & "Crescente"
            objComboOrdenacao.ItemData(objComboOrdenacao.NewIndex) = iIndice
            objComboOrdenacao.AddItem vColuna & SEPARADOR & "Decrescente"
            objComboOrdenacao.ItemData(objComboOrdenacao.NewIndex) = -iIndice
        End If
    Next
    
    Ordenacao_Preeenche = SUCESSO

    Exit Function

Erro_Ordenacao_Preeenche:

    Ordenacao_Preeenche = gErr

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182260)

    End Select

    Exit Function
    
End Function

Private Function SCodigo_Extrai(sString As String) As String

Dim iPosicaoSeparador As Integer, sBuffer As String
On Error GoTo Erro_SCodigo_Extrai

    iPosicaoSeparador = InStr(sString, SEPARADOR)
    If iPosicaoSeparador > 1 Then
        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
    Else
        sBuffer = sString
    End If
    
    SCodigo_Extrai = sBuffer

    Exit Function

Erro_SCodigo_Extrai:

    SCodigo_Extrai = ""

    Exit Function

End Function

Private Sub ListOuCombo_Seleciona_ItemData(objBox As Object, ByVal lItemData As Long)
'seleciona o item da combo ou listbox
Dim lErro As Long, iIndice As Integer
On Error GoTo Erro_ListOuCombo_Seleciona_ItemData

    'procura um item com o iTemData informado
    For iIndice = 0 To objBox.ListCount - 1

        If objBox.ItemData(iIndice) = lItemData Then
            objBox.ListIndex = iIndice
            Exit For
        End If
        
    Next

    Exit Sub

Erro_ListOuCombo_Seleciona_ItemData:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142348)

    End Select

    Exit Sub

End Sub

Public Function Soma_Coluna_Grid(ByVal objGridInt As AdmGrid, ByVal iColuna As Integer, ByVal objControle As Object, ByVal bReposicionaControle As Boolean, Optional ByVal iColSel As Integer = 0) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValorAcumulado As Double
Dim sString As String
Dim bConsidera As Boolean

On Error GoTo Erro_Soma_Coluna_Grid

    For iIndice = 1 To objGridInt.iLinhasExistentes
    
        bConsidera = True
        If iColSel <> 0 Then
            If objGridInt.objGrid.TextMatrix(iIndice, iColSel) = CStr(DESMARCADO) Then bConsidera = False
        End If
    
        If bConsidera Then
            sString = objGridInt.objGrid.TextMatrix(iIndice, iColuna)
            
            If Not IsNumeric(sString) And Len(Trim(sString)) > 0 Then gError 198553
        
            If Len(Trim(sString)) > 0 Then
                dValorAcumulado = dValorAcumulado + CDbl(sString)
            End If
        End If
    
    Next
    
    objControle.Caption = Format(dValorAcumulado, "STANDARD")
    
    If bReposicionaControle Then
    
        objControle.top = objGridInt.objGrid.top + objGridInt.objGrid.Height
        objControle.left = objGridInt.objGrid.left
        For iIndice = 0 To iColuna - 1
            objControle.left = objControle.left + objGridInt.objGrid.ColWidth(iIndice) + objGridInt.objGrid.GridLineWidth + 20
        Next
        objControle.Width = objGridInt.objGrid.ColWidth(iColuna)
    
    End If
    
    Soma_Coluna_Grid = SUCESSO
    
    Exit Function
    
Erro_Soma_Coluna_Grid:

    Soma_Coluna_Grid = gErr
    
    Select Case gErr
        
        Case 198553
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", gErr, sString)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198554)
        
    End Select

    Exit Function

End Function

Public Function Grid_Marca_Desmarca(ByVal objGridInt As AdmGrid, ByVal iColuna As Integer, ByVal iFlag As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Grid_Marca_Desmarca

    For iIndice = 1 To objGridInt.iLinhasExistentes
        objGridInt.objGrid.TextMatrix(iIndice, iColuna) = CStr(iFlag)
    Next
    
    Call Grid_Refresh_Checkbox(objGridInt)
    
    Grid_Marca_Desmarca = SUCESSO
    
    Exit Function
    
Erro_Grid_Marca_Desmarca:

    Grid_Marca_Desmarca = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198554)
        
    End Select

    Exit Function

End Function

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bIncluiLinha As Boolean = False) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim objCritica As New AdmCriticaExt

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = objCritica.Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bIncluiLinha Then
            lErro = Adiciona_Linha_Grid(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208635)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bIncluiLinha As Boolean = False) As Long
'Faz a crítica da célula Valor que está deixando de ser a corrente

Dim lErro As Long
Dim objCritica As New AdmCriticaExt

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = objCritica.Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objControle.Text = Format(objControle.Text, "STANDARD")

        If bIncluiLinha Then
            lErro = Adiciona_Linha_Grid(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208636)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

