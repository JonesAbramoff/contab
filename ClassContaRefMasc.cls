VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContaRefMasc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Inicializa_Mascara_ContaRef(ByVal objControle As Object) As Long
'inicializa a mascara da ContaRef

Dim sMascaraContaRef As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_ContaRef

    If Controle_ObterNomeClasse(objControle) = "MaskEdBox" Then

        'Inicializa a máscara da ContaRef
        sMascaraContaRef = String(STRING_CONTA, 0)
        
        'Armazena em sMascaraContaRef a mascara a ser a ser exibida no campo ContaRef
        lErro = MascaraItem(SEGMENTO_CONTAREF, sMascaraContaRef)
        If lErro <> SUCESSO Then gError 202010
        
        'coloca a mascara na tela.
        objControle.Mask = sMascaraContaRef
               
    End If
    
    Inicializa_Mascara_ContaRef = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_ContaRef:

    Inicializa_Mascara_ContaRef = gErr
    
    Select Case gErr
    
        Case 202010
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202011)
        
    End Select

    Exit Function

End Function

Public Function Retorna_ContaRef_Controle(ByVal objControle As Object, ByVal sValor As String) As Long
'inicializa a mascara da ContaRef

Dim sMascaraContaRef As String
Dim lErro As Long

On Error GoTo Erro_Retorna_ContaRef_Controle

    If Len(Trim(sValor)) <> 0 Then
    
        sMascaraContaRef = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_CONTAREF, sValor, sMascaraContaRef)
        If lErro <> SUCESSO Then gError 202012
    
        objControle.PromptInclude = False
        objControle.Text = sMascaraContaRef
        objControle.PromptInclude = True
        
    Else
    
        objControle.PromptInclude = False
        objControle.Text = ""
        objControle.PromptInclude = True
        
    End If
    
    Retorna_ContaRef_Controle = SUCESSO
    
    Exit Function
    
Erro_Retorna_ContaRef_Controle:

    Retorna_ContaRef_Controle = gErr
    
    Select Case gErr
    
        Case 202012
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202013)
        
    End Select

    Exit Function

End Function

Public Function Retorna_ContaRef_Tela(ByVal sValor As String, sValorMasc As String) As Long
'inicializa a mascara da ContaRef

Dim lErro As Long

On Error GoTo Erro_Retorna_ContaRef_Tela

    If Len(Trim(sValor)) <> 0 Then
    
        sValorMasc = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaItemTela(SEGMENTO_CONTAREF, sValor, sValorMasc)
        If lErro <> SUCESSO Then gError 202014
        
    Else
    
        sValorMasc = ""
        
    End If
    
    Retorna_ContaRef_Tela = SUCESSO
    
    Exit Function
    
Erro_Retorna_ContaRef_Tela:

    Retorna_ContaRef_Tela = gErr
    
    Select Case gErr
    
        Case 202014
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202015)
        
    End Select

    Exit Function

End Function

Public Function Retorna_ContaRef_Enxuto(ByVal sValor As String, sValorMasc As String) As Long
'inicializa a mascara da ContaRef

Dim lErro As Long

On Error GoTo Erro_Retorna_ContaRef_Enxuto

    If Len(Trim(sValor)) <> 0 Then
    
        sValorMasc = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_CONTAREF, sValor, sValorMasc)
        If lErro <> SUCESSO Then gError 202016
        
    Else
    
        sValorMasc = ""
        
    End If
    
    Retorna_ContaRef_Enxuto = SUCESSO
    
    Exit Function
    
Erro_Retorna_ContaRef_Enxuto:

    Retorna_ContaRef_Enxuto = gErr
    
    Select Case gErr
    
        Case 202016
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202017)
        
    End Select

    Exit Function

End Function

Public Function Mascara_Mascarar_ContaRef(ByVal sValor As String, sValorMasc As String) As Long
'inicializa a mascara da ContaRef

Dim lErro As Long

On Error GoTo Erro_Mascara_Mascarar_ContaRef

    If Len(Trim(sValor)) <> 0 Then
    
        sValorMasc = String(STRING_CONTA, 0)
        
        Select Case Len(Trim(sValor))
        
            Case 1
                sValorMasc = left(sValor, 1)
            Case 3
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2)
            Case 5
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2) & "." & Mid(sValor, 4, 2)
            Case 7
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2) & "." & Mid(sValor, 4, 2) & "." & Mid(sValor, 6, 2)
            Case 9
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2) & "." & Mid(sValor, 4, 2) & "." & Mid(sValor, 6, 2) & "." & Mid(sValor, 8, 2)
            Case 11
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2) & "." & Mid(sValor, 4, 2) & "." & Mid(sValor, 6, 2) & "." & Mid(sValor, 8, 2) & "." & Mid(sValor, 10, 2)
            Case 13
                sValorMasc = left(sValor, 1) & "." & Mid(sValor, 2, 2) & "." & Mid(sValor, 4, 2) & "." & Mid(sValor, 6, 2) & "." & Mid(sValor, 8, 2) & "." & Mid(sValor, 10, 2) & "." & Mid(sValor, 12, 2)
        
        End Select
    
        'lErro = Mascara_MascararItem(SEGMENTO_CONTAREF, sValor, sValorMasc)
        If lErro <> SUCESSO Then gError 202018
        
    Else
    
        sValorMasc = ""
        
    End If
    
    Mascara_Mascarar_ContaRef = SUCESSO
    
    Exit Function
    
Erro_Mascara_Mascarar_ContaRef:

    Mascara_Mascarar_ContaRef = gErr
    
    Select Case gErr
    
        Case 202018
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202019)
        
    End Select

    Exit Function

End Function

Public Function ContaRef_Formata(ByVal sValor As String, sValorFormatado As String, iCampoPreenchido As Integer) As Long
'inicializa a mascara da ContaRef

Dim lErro As Long

On Error GoTo Erro_ContaRef_Formata

    sValorFormatado = String(STRING_CONTA, 0)
    
    
    sValorFormatado = Replace(sValor, ".", "")
    sValorFormatado = sValorFormatado & String(13 - Len(sValorFormatado), " ")
    
    'Coloca no formato do BD
    'lErro = CF("Item_Formata", SEGMENTO_CONTAREF, sValor, sValorFormatado, iCampoPreenchido)
    If lErro <> SUCESSO Then gError 202020
    
    ContaRef_Formata = SUCESSO
    
    Exit Function
    
Erro_ContaRef_Formata:

    ContaRef_Formata = gErr
    
    Select Case gErr
    
        Case 202020
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202021)
        
    End Select

    Exit Function

End Function

Public Function Retorna_ContaRef_Pai(ByVal sConta As String, sContaPai As String) As Long
'inicializa a mascara da ContaRef

Dim sMascaraContaRef As String
Dim lErro As Long

On Error GoTo Erro_Retorna_ContaRef_Pai

    If Len(Trim(sConta)) <> 0 Then
    
        sContaPai = String(STRING_CONTA, 0)
        
            
        Select Case Len(Trim(sConta))
        
            Case 1
                sContaPai = ""
            Case 3
                sContaPai = left(sConta, 1) & String(12, " ")
            Case 5
                sContaPai = left(sConta, 3) & String(10, " ")
            Case 7
                sContaPai = left(sConta, 5) & String(8, " ")
            Case 9
                sContaPai = left(sConta, 7) & String(6, " ")
            Case 11
                sContaPai = left(sConta, 9) & String(4, " ")
            Case 13
                sContaPai = left(sConta, 11) & String(2, " ")
        
        End Select
    
        'lErro = Mascara_RetornaItemPai(SEGMENTO_CONTAREF, sConta, sContaPai)
        If lErro <> SUCESSO Then gError 202022
           
    Else
    
        sConta = ""
        
    End If
    
    Retorna_ContaRef_Pai = SUCESSO
    
    Exit Function
    
Erro_Retorna_ContaRef_Pai:

    Retorna_ContaRef_Pai = gErr
    
    Select Case gErr
    
        Case 202022
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202023)
        
    End Select

    Exit Function

End Function

Public Function ContaRef_Nivel(ByVal sValor As String, iNivel As Integer) As Long
'inicializa a mascara da ContaRef

Dim lErro As Long

On Error GoTo Erro_ContaRef_Nivel

    lErro = Mascara_Item_ObterNivel(SEGMENTO_CONTAREF, sValor, iNivel)
    If lErro <> SUCESSO Then gError 202024
    
    ContaRef_Nivel = SUCESSO
    
    Exit Function
    
Erro_ContaRef_Nivel:

    ContaRef_Nivel = gErr
    
    Select Case gErr
    
        Case 202024
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202025)
        
    End Select

    Exit Function

End Function



