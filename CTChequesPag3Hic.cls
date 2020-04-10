VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTChequesPag3Hic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Observacao(ByVal objCT As Object) As Object
     Set Observacao = objCT.objUserControl.Controls("Observacao")
End Property

Public Sub Observacao_GotFocus(ByVal objCT As Object)
    
    Call Grid_Campo_Recebe_Foco(objCT.objGridChequesPag3)

End Sub

Public Sub Observacao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridChequesPag3)
    
End Sub

Public Sub Observacao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridChequesPag3.objControle = Observacao(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function CHQ_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula do Grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_CHQ_Saida_Celula

    If objGridInt.objGrid Is objCT.GridChequesPag3 Then

        Select Case objGridInt.objGrid.Col
        
            'Se a célula for o campo Cheque
            Case objCT.iGrid_Beneficiario_Col
                
                Set objGridInt.objControle = objCT.Beneficiario
                
               'Chama função de tratamento de saída da célula Beneficiário
                lErro = Saida_Celula_Beneficiario(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132561
                
            Case objCT.gobjInfoUsu.iGrid_Observacao_Col

                Set objGridInt.objControle = Observacao(objCT)
                
               'Chama função de tratamento de saída da célula Beneficiário
                lErro = Saida_Celula_Observacao(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 132562
                
        End Select

    End If

    CHQ_Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_CHQ_Saida_Celula:

    CHQ_Saida_Celula = gErr
    
    Select Case gErr
    
        Case 132561, 132562

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179248)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Beneficiario(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_Saida_Celula_Beneficiario

    Set objGridInt.objControle = objCT.Beneficiario

    If Not (objCT.gobjChequesPag Is Nothing) Then
    
        'Passa os dados da linha do Grid para o Obj
        Set objInfoChequePag = objCT.gobjChequesPag.ColInfoChequePag.Item(objCT.GridChequesPag3.Row)
            
        'Passa para o Obj o valor do Cheque que está na tela
        objInfoChequePag.sFavorecido = objCT.Beneficiario.Text
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132562
    
    Saida_Celula_Beneficiario = SUCESSO

    Exit Function

Erro_Saida_Celula_Beneficiario:

    Saida_Celula_Beneficiario = gErr

    Select Case gErr

        Case 132562
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179249)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Observacao(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao(objCT)

    If Not (objCT.gobjChequesPag Is Nothing) Then
    
        'Passa os dados da linha do Grid para o Obj
        Set objInfoChequePag = objCT.gobjChequesPag.ColInfoChequePag.Item(objCT.GridChequesPag3.Row)
            
        'Passa para o Obj o valor do Cheque que está na tela
        objInfoChequePag.sObservacao = Observacao(objCT).Text
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132563
    
    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 132563
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179250)

    End Select

    Exit Function

End Function

Public Function CHQ_Inicializa_Grid_ChequesPag3(ByVal objCT As Object, objGridInt As AdmGrid, iProximo As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CHQ_Inicializa_Grid_ChequesPag3
    
    'titulos do grid
    objGridInt.colColuna.Add ("Observacao")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Observacao(objCT).Name)
        
    objCT.gobjInfoUsu.iGrid_Observacao_Col = iProximo

    iProximo = iProximo + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    CHQ_Inicializa_Grid_ChequesPag3 = SUCESSO
    
    Exit Function
    
Erro_CHQ_Inicializa_Grid_ChequesPag3:

    CHQ_Inicializa_Grid_ChequesPag3 = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179251)
        
    End Select

    Exit Function
        
End Function

Public Function CHQ_Trata_Parametros(ByVal objCT As Object, ByVal objInfoChequePag As ClassInfoChequePag, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CHQ_Trata_Parametros
    
    objCT.GridChequesPag3.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Observacao_Col) = objInfoChequePag.sObservacao

    CHQ_Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_CHQ_Trata_Parametros:

    CHQ_Trata_Parametros = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179252)
        
    End Select

    Exit Function
        
End Function

Public Sub BotaoReter_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_BotaoReter_Click

   lErro = CF("ParcelaPag_Grava_Cheque", objCT.gobjChequesPag)
   If lErro <> SUCESSO Then gError 132582
   
   Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUE_RETIDO")

   Exit Sub

Erro_BotaoReter_Click:

    Select Case gErr
    
        Case 132582
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179253)
    
    End Select

    Exit Sub
    
End Sub

