VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArvore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Carga_Arvore_Conta_Modulo(ByVal colNodes As Nodes, ByVal sSiglaModulo As String) As Long
'move os dados do plano de contas, relativos ao modulo em questão,  do banco de dados para a arvore colNodes.

Dim colPlanoConta As New Collection
Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim sContaPai As String
Dim objNode As Node

On Error GoTo Erro_Carga_Arvore_Conta_Modulo

    'le todas as contas de nível 0 e 1 do modulo em questao e coloca-as em colPlanoConta
    lErro = CF("PlanoConta_Le_Niveis0e1_Modulo", colPlanoConta, sSiglaModulo)
    If lErro <> SUCESSO Then Error 40802
    
     For Each objPlanoConta In colPlanoConta

        sContaMascarada = String(STRING_CONTA, 0)

        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 40803

        If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
            sConta = "A" & objPlanoConta.sConta
        Else
            sConta = "S" & objPlanoConta.sConta
        End If

        sContaPai = String(STRING_CONTA, 0)

        'retorna a conta "pai" da conta em questão, se houver
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then Error 40804

        'se a conta possui uma conta "pai"
        If Len(Trim(sContaPai)) > 0 Then

            sContaPai = "S" & sContaPai

            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)

        Else
            'se a conta não possui conta "pai"
            Set objNode = colNodes.Add(, tvwLast, sConta)

        End If

        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta

    Next

    Carga_Arvore_Conta_Modulo = SUCESSO

    Exit Function

Erro_Carga_Arvore_Conta_Modulo:

    Carga_Arvore_Conta_Modulo = Err

    Select Case Err

        Case 40802

        Case 40803
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)

        Case 40804
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144637)

    End Select

    Exit Function

End Function
    

'Private Function Carga_No_Arvore(colNodes As Nodes, objConta As ClassPlanoConta, sContaAntiga As String, colSinteticas As Collection) As Long
''verifica se a conta no nivel passado como parametro já está na arvore. Se não estiver coloca
'
'Dim objNode As Node
'Dim sContaPai As String
'Dim sContaAntigaPai As String
'Dim sConta1 As String
'Dim sContaMascarada As String
'Dim objContaPai As ClassPlanoConta
'Dim objContaAntigaPai As ClassPlanoConta
'Dim lErro As Long
'
'
'    If objConta.iNivelConta > 1 Then
'
'        sContaPai = String(STRING_CONTA, 0)
'        sContaAntigaPai = String(STRING_CONTA, 0)
'
'        lErro = Mascara_RetornaContaPai(objConta.sConta, sContaPai)
'        If lErro <> SUCESSO Then Error 10062
'
'        lErro = Mascara_RetornaContaPai(sContaAntiga, sContaAntigaPai)
'        If lErro <> SUCESSO Then Error 10063
'
'        If sContaPai <> sContaAntigaPai Then
'
'            Set objContaPai = colSinteticas.Item(sContaPai)
'
'            lErro = Carga_No_Arvore(colNodes, objContaPai, sContaAntigaPai, colSinteticas)
'            If lErro <> SUCESSO Then Error 10064
'
'        End If
'
'    End If
'
'
'    If iTipoConta = CONTA_ANALITICA Then
'        sConta = "A" & objPlanoConta.sConta
'    Else
'        sConta = "S" & objPlanoConta.sConta
'    End If
'
'    'se a conta possui uma conta "pai"
'    If iNivelConta > 1 Then
'
'        sContaPai = "S" & sContaPai
'
'        Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)
'
'    Else
'        'se a conta não possui conta "pai"
'        Set objNode = colNodes.Add(, tvwLast, sConta)
'
'    End If
'
'    sContaMascarada = String(STRING_CONTA, 0)
'
'    'coloca a conta no formato que é exibida na tela
'    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
'    If lErro <> SUCESSO Then Error 10065
'
'    objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
'
'End Function
    
Private Function Carga_No_Sintetico_Arvore(ByVal colNodes As Nodes, ByVal objPlanoConta As ClassPlanoConta, ByVal colSinteticas As Collection, ByVal colAnaliticas As Collection, ByVal iNivelConta As Integer, iIndice As Integer) As Long
'verifica se a conta no nivel passado como parametro já está na arvore. Se não estiver coloca

Dim objNode As Node
Dim sContaAtual As String
Dim sContaPai As String
Dim sContaAntiga As String
Dim sContaMascarada As String
Dim objPlanoContaAtual As ClassPlanoConta
Dim sConta1 As String
Dim lErro As Long

On Error GoTo Erro_Carga_No_Sintetico_Arvore

    sContaAntiga = ""

    Do While colAnaliticas.Count >= iIndice And objPlanoConta.iNivelConta >= iNivelConta

        sContaAtual = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaNoNivel(iNivelConta, objPlanoConta.sConta, sContaAtual)
        If lErro <> SUCESSO Then Error 10062
        
        If sContaAtual = objPlanoConta.sConta Then
            
            lErro = Carga_No_Analitico_Arvore(colNodes, objPlanoConta, colAnaliticas, iIndice)
            If lErro <> SUCESSO Then Error 10066
            
            Exit Do
            
        Else
        
            If sContaAtual <> sContaAntiga Then
            
                sContaAntiga = sContaAtual
        
                Set objPlanoContaAtual = colSinteticas.Item(sContaAtual)
            
                sConta1 = "S" & objPlanoContaAtual.sConta
    
                'se a conta possui uma conta "pai"
                If iNivelConta > 1 Then
                
                    sContaPai = String(STRING_CONTA, 0)
    
                    lErro = Mascara_RetornaContaPai(sContaAtual, sContaPai)
                    If lErro <> SUCESSO Then Error 10068
    
                    sContaPai = "S" & sContaPai
            
                    Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta1)
    
                Else
                
                    'se a conta não possui conta "pai"
                    Set objNode = colNodes.Add(, tvwLast, sConta1)
            
                End If
        
                sContaMascarada = String(STRING_CONTA, 0)
            
                'coloca a conta no formato que é exibida na tela
                lErro = Mascara_MascararConta(sContaAtual, sContaMascarada)
                If lErro <> SUCESSO Then Error 10065
        
                objNode.Text = sContaMascarada & SEPARADOR & objPlanoContaAtual.sDescConta
                
            End If
            
            lErro = Carga_No_Sintetico_Arvore(colNodes, objPlanoConta, colSinteticas, colAnaliticas, iNivelConta + 1, iIndice)
            If lErro <> SUCESSO Then Error 10067
            
            If iNivelConta = objPlanoConta.iNivelConta Then Exit Do
            
        End If
            
   Loop
            
   Carga_No_Sintetico_Arvore = SUCESSO
   
   Exit Function
   
Erro_Carga_No_Sintetico_Arvore:

    Carga_No_Sintetico_Arvore = Err
    
    Select Case Err

        Case 10062
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, objPlanoConta.sConta, iNivelConta)
        
        Case 10065
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, sContaAtual)
        
        Case 10066, 10067
        
        Case 10068
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, sContaAtual)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144638)

    End Select
    
    Exit Function
            
End Function

Private Function Carga_No_Analitico_Arvore(ByVal colNodes As Nodes, ByVal objPlanoConta As ClassPlanoConta, ByVal colAnaliticas As Collection, iIndice As Integer) As Long
'Coloca a conta analitica na arvore e acessa a seguinte

Dim objNode As Node
Dim sContaPai As String
Dim sContaMascarada As String
Dim sConta1 As String
Dim lErro As Long

On Error GoTo Erro_Carga_No_Analitico_Arvore

    sConta1 = "A" & objPlanoConta.sConta

    'se a conta possui uma conta "pai"
    If objPlanoConta.iNivelConta > 1 Then
    
        sContaPai = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then Error 10069

        sContaPai = "S" & sContaPai

        Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta1)

    Else
    
        'se a conta não possui conta "pai"
        Set objNode = colNodes.Add(, tvwLast, sConta1)

    End If

    sContaMascarada = String(STRING_CONTA, 0)

    'coloca a conta no formato que é exibida na tela
    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
    If lErro <> SUCESSO Then Error 10070

    objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
    
    iIndice = iIndice + 1
    
    If colAnaliticas.Count >= iIndice Then
        Set objPlanoConta = colAnaliticas.Item(iIndice)
    End If
                        
   Carga_No_Analitico_Arvore = SUCESSO
   
   Exit Function
   
Erro_Carga_No_Analitico_Arvore:

    Carga_No_Analitico_Arvore = Err
    
    Select Case Err

        Case 10069
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)
        
        Case 10070
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144639)

    End Select
    
    Exit Function
            
End Function

Function Carga_Arvore_Conta_Modulo1(ByVal objNodeAvo As Node, ByVal colNodes As Nodes, ByVal sSiglaModulo As String) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colPlanoConta As New Collection
Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim sContaAvo As String
Dim sContaPai As String
    
On Error GoTo Erro_Carga_Arvore_Conta_Modulo1
    
    sContaAvo = Mid(objNodeAvo.Key, 2)
    
    'le os filhos da conta em questão e coloca-as em colPlanoConta
    lErro = CF("PlanoConta_Le_Netos_Modulo", sContaAvo, colPlanoConta, sSiglaModulo)
    If lErro <> SUCESSO Then Error 40805
    
    For Each objPlanoConta In colPlanoConta
    
        sContaMascarada = String(STRING_CONTA, 0)
        
        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 40806

        If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
            sConta = "A" & objPlanoConta.sConta
        Else
            sConta = "S" & objPlanoConta.sConta
        End If

        sContaPai = String(STRING_CONTA, 0)
        
        'retorna a conta "pai" da conta em questão, se houver
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then Error 40807
        
        'se a conta possui uma conta "pai"
        If Len(Trim(sContaPai)) > 0 Then

            sContaPai = "S" & sContaPai
            
            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)

        Else
            'se a conta não possui conta "pai"
            Set objNode = colNodes.Add(, tvwLast, sConta)
            
        End If
        
        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Conta_Modulo1 = SUCESSO

    Exit Function

Erro_Carga_Arvore_Conta_Modulo1:

    Carga_Arvore_Conta_Modulo1 = Err

    Select Case Err

        Case 40805
        
        Case 40806
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
            
        Case 40807
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144640)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Conta(ByVal colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colPlanoConta As New Collection
Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim sContaPai As String
    
On Error GoTo Erro_Carga_Arvore_Conta
    
    'le todas as contas de nível 0 e 1 e coloca-as em colPlanoConta
    lErro = CF("PlanoConta_Le_Niveis0e1", colPlanoConta)
    If lErro <> SUCESSO Then Error 5911
    
    For Each objPlanoConta In colPlanoConta
    
        sContaMascarada = String(STRING_CONTA, 0)
        
        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 5912

        If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
            sConta = "A" & objPlanoConta.sConta
        Else
            sConta = "S" & objPlanoConta.sConta
        End If

        sContaPai = String(STRING_CONTA, 0)
        
        'retorna a conta "pai" da conta em questão, se houver
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then Error 5913
        
        'se a conta possui uma conta "pai"
        If Len(Trim(sContaPai)) > 0 Then

            sContaPai = "S" & sContaPai
            
            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)

        Else
            'se a conta não possui conta "pai"
            Set objNode = colNodes.Add(, tvwLast, sConta)
            
        End If
        
        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
    
    Next
    
    Carga_Arvore_Conta = SUCESSO

    Exit Function

Erro_Carga_Arvore_Conta:

    Carga_Arvore_Conta = Err

    Select Case Err

        Case 5911
        
        Case 5912
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
            
        Case 5913
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144641)

    End Select
    
    Exit Function

End Function


Function Carga_Arvore_Conta1(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colPlanoConta As New Collection
Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim sContaAvo As String
Dim sContaPai As String
    
    
On Error GoTo Erro_Carga_Arvore_Conta1
    
    sContaAvo = Mid(objNodeAvo.Key, 2)
    
    'le os filhos da conta em questão e coloca-as em colPlanoConta
    lErro = CF("PlanoConta_Le_Netos", sContaAvo, colPlanoConta)
    If lErro <> SUCESSO Then Error 44024
    
    For Each objPlanoConta In colPlanoConta
    
        sContaMascarada = String(STRING_CONTA, 0)
        
        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 44025

        If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
            sConta = "A" & objPlanoConta.sConta
        Else
            sConta = "S" & objPlanoConta.sConta
        End If

        sContaPai = String(STRING_CONTA, 0)
        
        'retorna a conta "pai" da conta em questão, se houver
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then Error 44026
        
        'se a conta possui uma conta "pai"
        If Len(Trim(sContaPai)) > 0 Then

            sContaPai = "S" & sContaPai
            
            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)

        Else
            'se a conta não possui conta "pai"
            Set objNode = colNodes.Add(, tvwLast, sConta)
            
        End If
        
        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Conta1 = SUCESSO

    Exit Function

Erro_Carga_Arvore_Conta1:

    Carga_Arvore_Conta1 = Err

    Select Case Err

        Case 44024
        
        Case 44025
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
            
        Case 44026
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144642)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Ccl(ByVal colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes. /m

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    'leitura dos centro de custo/lucro no BD
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then gError 178932
    
    'para cada centro de custo encontrado no bd
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        'coloca a mascara no centro de custo
        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 178933

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then gError 178934
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "S" & sCclPai
            
            'adiciona o centro de custo como filho do centro de custo pai
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
        
            'se o centro de custo/lucro não possui centro de custo/lucro "pai", adiciona na árvore sem pai
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        'coloca o texto do nó que acabou de ser inserido
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = gErr

    Select Case gErr

        Case 178932

        Case 178933
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objCcl.sCcl)

        Case 178934
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178935)

    End Select
    
    Exit Function

End Function

