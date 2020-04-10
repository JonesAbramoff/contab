VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmArvore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Carga_Arvore_Produto(ByVal colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
    
On Error GoTo Erro_Carga_Arvore_Produto
    
    'le todos produtos de nível 0 e 1 e coloca-os em colProduto
    lErro = CF("Produto_Le_Niveis0e1",colProduto)
    If lErro <> SUCESSO Then Error 44130
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 44131

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 44132

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    Carga_Arvore_Produto = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto:

    Carga_Arvore_Produto = Err

    Select Case Err

        Case 44130
        
        Case 44131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 44132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142042)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Reserva(ByVal colNodes As Nodes) As Long
'Move os dados de Produtos do BD que têm reserva e são ativos para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String

On Error GoTo Erro_Carga_Arvore_Produto_Reserva

    'leitura das Produtos no BD
    lErro = CF("Produto_Le_Niveis0e1_Reserva",colProduto)
    If lErro <> SUCESSO Then Error 44160

    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 44161

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 44162

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    Carga_Arvore_Produto_Reserva = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Reserva:

    Carga_Arvore_Produto_Reserva = Err

    Select Case Err

        Case 44160

        Case 44161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 44162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142043)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Netos(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String
    
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos
    
    sProdutoAvo = Mid(objNodeAvo.Key, 2)
    
    'le os netos do produto em questão e coloca-os em colProduto
    lErro = CF("Produto_Le_Netos",sProdutoAvo, colProduto)
    If lErro <> SUCESSO Then Error 44139
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 44140

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 44141

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Produto_Netos = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos:

    Carga_Arvore_Produto_Netos = Err

    Select Case Err

        Case 44139
        
        Case 44140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 44141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142044)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Netos_Reserva(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String
    
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Reserva
    
    sProdutoAvo = Mid(objNodeAvo.Key, 2)
    
    'le os netos do produto em questão e coloca-os em colProduto
    lErro = CF("Produto_Le_Netos_Reserva",sProdutoAvo, colProduto)
    If lErro <> SUCESSO Then Error 44163
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 44164

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 44165

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Produto_Netos_Reserva = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos_Reserva:

    Carga_Arvore_Produto_Netos_Reserva = Err

    Select Case Err

        Case 44163
        
        Case 44164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 44165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142045)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Componentes(ByVal colNodes As Nodes) As Long
'Move os dados de Produtos do BD que podem participar de kit e são ativos para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String

On Error GoTo Erro_Carga_Arvore_Produto_Componentes

    'leitura das Produtos no BD
    lErro = CF("Produto_Le_Niveis0e1_Componentes",colProduto)
    If lErro <> SUCESSO Then Error 49957

    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 49958

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 49959

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    Carga_Arvore_Produto_Componentes = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Componentes:

    Carga_Arvore_Produto_Componentes = Err

    Select Case Err

        Case 49957

        Case 49958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 49959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142046)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Netos_Componentes(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String
    
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Componentes
    
    sProdutoAvo = Mid(objNodeAvo.Key, 2)
    
    'le os netos do produto em questão e coloca-os em colProduto
    lErro = CF("Produto_Le_Netos_Componentes",sProdutoAvo, colProduto)
    If lErro <> SUCESSO Then Error 49966
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 49967

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 49968

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Produto_Netos_Componentes = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos_Componentes:

    Carga_Arvore_Produto_Netos_Componentes = Err

    Select Case Err

        Case 49966
        
        Case 49967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 49968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142047)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Netos_Venda(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Venda
    
    sProdutoAvo = Mid(objNodeAvo.Key, 2)
    
    If giFilialEmpresa = EMPRESA_TODA Then
        'le os netos do produto em questão e coloca-os em colProduto
        lErro = CF("Produto_Le_Netos_Venda_EmpresaToda",sProdutoAvo, colProduto)
        If lErro <> SUCESSO Then Error 61182
    Else
        'le os netos do produto em questão e coloca-os em colProduto
        lErro = CF("Produto_Le_Netos_Venda",sProdutoAvo, colProduto)
        If lErro <> SUCESSO Then Error 58058
    End If
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 58059

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 58060

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Produto_Netos_Venda = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos_Venda:

    Carga_Arvore_Produto_Netos_Venda = Err

    Select Case Err

        Case 58058, 61182
        
        Case 58059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 58060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142048)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Venda(ByVal colNodes As Nodes) As Long
'Move os dados de Produtos do BD que pode ser Vendido e são ativos para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String

On Error GoTo Erro_Carga_Arvore_Produto_Venda

    If giFilialEmpresa = EMPRESA_TODA Then
        'Leitura dos Produtos que Podem ser vendidos no BD para Empresa Toda
        lErro = CF("Produto_Le_Niveis0e1_Venda_EmpresaToda",colProduto)
        If lErro <> SUCESSO Then Error 61181
    Else
        'Leitura dos Produtos que Podem ser vendidos no BD por FilialEmpresa
        lErro = CF("Produto_Le_Niveis0e1_Venda",colProduto)
        If lErro <> SUCESSO Then Error 58055
    End If
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 58056

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 58057

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    Carga_Arvore_Produto_Venda = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Venda:

    Carga_Arvore_Produto_Venda = Err

    Select Case Err

        Case 58055, 61181

        Case 58056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 58057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142049)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Inventariado(ByVal colNodes As Nodes) As Long
'Move os dados de Produtos do BD que têm Produto Inventariado e são ativos para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String

On Error GoTo Erro_Carga_Arvore_Produto_Inventariado

    'leitura das Produtos no BD
    lErro = CF("Produto_Le_Niveis0e1_Inventariado",colProduto)
    If lErro <> SUCESSO And lErro <> 52138 Then Error 52140
    
    If lErro = 52138 Then Error 52152
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 52141

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 52142

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    Carga_Arvore_Produto_Inventariado = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Inventariado:

    Carga_Arvore_Produto_Inventariado = Err

    Select Case Err

        Case 52140

        Case 52141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 52142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case 52152 ' ( Caso não tenha encontrado nenhum Produto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142050)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Netos_Inventariado(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String
    
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Inventariado
    
    sProdutoAvo = Mid(objNodeAvo.Key, 2)
    
    'le os netos do produto em questão e coloca-os em colProduto
    lErro = CF("Produto_Le_Netos_Inventariado",sProdutoAvo, colProduto)
    If lErro <> SUCESSO Then Error 52143
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 52144

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 52145

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next
    
    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE
    
    Carga_Arvore_Produto_Netos_Inventariado = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos_Inventariado:

    Carga_Arvore_Produto_Netos_Inventariado = Err

    Select Case Err

        Case 52143
        
        Case 52144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 52145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142051)

    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Estoque(ByVal colNodes As Nodes) As Long
'Move os dados de Produtos do BD que são Estocaveis e são ativos para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String

On Error GoTo Erro_Carga_Arvore_Produto_Estoque

    'leitura das Produtos no BD
    lErro = CF("Produto_Le_Niveis0e1_Estoque",colProduto)
    If lErro <> SUCESSO And lErro <> 58753 Then Error 58761

    If lErro = 58753 Then Error 58762

    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 58763

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 58764

        'se a Produto possui uma Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    Carga_Arvore_Produto_Estoque = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Estoque:

    Carga_Arvore_Produto_Estoque = Err

    Select Case Err

        Case 58761

        Case 58762 ' ( Caso não tenha encontrado nenhum Produto)

        Case 58763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 58764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142052)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Netos_Estoque(ByVal objNodeAvo As Node, ByVal colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
Dim sProdutoAvo As String


On Error GoTo Erro_Carga_Arvore_Produto_Netos_Estoque

    sProdutoAvo = Mid(objNodeAvo.Key, 2)

    'le os netos do produto em questão e coloca-os em colProduto
    lErro = CF("Produto_Le_Netos_Estoque",sProdutoAvo, colProduto)
    If lErro <> SUCESSO Then Error 58765

    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 58766

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 58767

        'se a Produto possui um Produto "pai"
        If Len(Trim(sProdutoPai)) > 0 Then

            sProdutoPai = "X" & sProdutoPai

            Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)

        Else
            'se a Produto não possui Produto "pai"
            Set objNode = colNodes.Add(, tvwLast, sProduto)

        End If

        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao

    Next

    'coloca o tag indicando que os netos já foram carregados
    objNodeAvo.Tag = NETOS_NA_ARVORE

    Carga_Arvore_Produto_Netos_Estoque = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Netos_Estoque:

    Carga_Arvore_Produto_Netos_Estoque = Err

    Select Case Err

        Case 58765

        Case 58766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 58767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142053)

    End Select

    Exit Function

End Function

Function Carga_Arvore_Produto_Produziveis(ByVal colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colProduto As New Collection
Dim objProduto As ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProduto As String
Dim sProdutoPai As String
    
On Error GoTo Erro_Carga_Arvore_Produto_Produziveis
    
    'le todos produtos de nível 0 e 1 e coloca-os em colProduto
    lErro = CF("Produto_Le_Niveis0e1",colProduto)
    If lErro <> SUCESSO Then Error 44130
    
    For Each objProduto In colProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 44131

        sProduto = "X" & objProduto.sCodigo

        sProdutoPai = String(STRING_PRODUTO, 0)

        'retorna a Produto "pai" da Produto em questão, se houver
        lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then Error 44132

        'Verifica se o produto é comprado. Se for, não inclui na lista.
        If Not ((objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL) And objProduto.iGerencial <> PRODUTO_GERENCIAL) Then

            'se a Produto possui uma Produto "pai"
            If Len(Trim(sProdutoPai)) > 0 Then
    
                sProdutoPai = "X" & sProdutoPai
    
                Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto)
    
            Else
                'se a Produto não possui Produto "pai"
                Set objNode = colNodes.Add(, tvwLast, sProduto)
    
            End If

            objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao
        
        End If
        
    Next
    
    Carga_Arvore_Produto_Produziveis = SUCESSO

    Exit Function

Erro_Carga_Arvore_Produto_Produziveis:

    Carga_Arvore_Produto_Produziveis = Err

    Select Case Err

        Case 44130
        
        Case 44131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)

        Case 44132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142054)

    End Select
    
    Exit Function

End Function

