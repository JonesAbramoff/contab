VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOMBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PedComprasVSReq_BotaoEdita(objPedCompraReq As ClassRequisicaoCompras, lErro As Long) As Long

On Error GoTo Erro_PedComprasVSReq_BotaoEdita

    'Se lErro retornar sucesso ----> Executa uma das condições abaixo
    If lErro = SUCESSO Then
    
        'Verifica DataEnvio recebida com parametro da função
        If objPedCompraReq.dtDataEnvio <> DATA_NULA Then

            Call Chama_Tela("ReqComprasEnv", objPedCompraReq)
                   
        'Caso as condições sejam Falsas ---> executa código abaixo
        Else
            
            Call Chama_Tela("ReqCompras", objPedCompraReq)
    
        End If
        
    End If
  
    PedComprasVSReq_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_PedComprasVSReq_BotaoEdita:

    PedComprasVSReq_BotaoEdita = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147431)

    End Select
 
    Exit Function
        
End Function

Function PedComprasItemReqLista_BotaoEdita(objPedidoCompras As ClassPedidoCompras, lErro As Long) As Long
'Rotina especial para tratamento do Botao Editar da Tela de Browser
'PedComprasItemReqLista

On Error GoTo Erro_PedComprasItemReqLista_BotaoEdita

    'Se lErro retornar sucesso ----> Executa uma das condições abaixo
    If lErro = SUCESSO Then
    
        'Verifica variável Status recebida com parametro da função
        If objPedidoCompras.iStatus = 1 Then
            Call Chama_Tela("PedidoCompraBaixado", objPedidoCompras)
        
        'Verifica DataEnvio recebida com parametro da função
        ElseIf objPedidoCompras.dtDataEnvio <> DATA_NULA Then
            Call Chama_Tela("PedComprasCons", objPedidoCompras)
            
        'Caso as condições sejam Falsas ---> executa código abaixo
        Else
            Call Chama_Tela("PedComprasGerado", objPedidoCompras)
    
        End If
    
    'Se lErro retornar erro ----> Executa Código abaixo
    Else
        Call Chama_Tela("PedidoCompras")
    
    End If
    
    'Função executada com Sucesso
    PedComprasItemReqLista_BotaoEdita = SUCESSO
    
    Exit Function

'--->Aqui se inicia o Tratamento dos Erros
Erro_PedComprasItemReqLista_BotaoEdita:

    PedComprasItemReqLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147432)

    End Select

    Exit Function

End Function

Function PedComprasNaoEnvLista_BotaoEdita(objPedidoCompra As ClassPedidoCompras, lErro As Long) As Long
'Rotina especial para tratamento do Botao Editar da Tela de Browser
'PedComprasNaoEnvLista
 
Dim objPedidoCompra1 As New ClassPedidoCompras
Dim iPedidoGerado As Integer

On Error GoTo Erro_PedComprasNaoEnvLista_BotaoEdita

    'Verifica variável Status recebida com parametro da função
    If lErro = SUCESSO Then
        
        'Atribui Valor de objPedidoCompra.lNumIntDoc para objPedidoCompra1.lNumIntDoc
        objPedidoCompra1.lNumIntDoc = objPedidoCompra.lNumIntDoc
        
        'Executa leitura na tabela
        lErro = CF("PedCompraGerado_Le", objPedidoCompra1)
        If lErro <> SUCESSO And lErro <> 89947 Then gError 89949

        'Se Retornou Sucesso ---> Seta Flag iPedidoGerado para 1
        If lErro = SUCESSO Then iPedidoGerado = 1

    End If

    'Se a condição for verdadeira ---> chama PedidoCompras
    If iPedidoGerado = 0 Then
        Call Chama_Tela("PedidoCompras", objPedidoCompra)
    
    'Se a condição for falsa ---> chama PedComprasGerado
    Else
        Call Chama_Tela("PedComprasGerado", objPedidoCompra)
    End If
    
    'Função executada com Sucesso
    PedComprasNaoEnvLista_BotaoEdita = SUCESSO
    
    Exit Function

'--->Aqui se inicia o Tratamento dos Erros
Erro_PedComprasNaoEnvLista_BotaoEdita:

    PedComprasNaoEnvLista_BotaoEdita = gErr

    Select Case gErr
    
        Case 89949
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147433)

    End Select

    Exit Function

End Function

Function PedidoCompraAbertoLista_BotaoEdita(objPedidoCompras As ClassPedidoCompras, lErro As Long) As Long
'Rotina especial para tratamento do Botao Editar da Tela de Browser
'PedidoCompraAbertoLista

On Error GoTo Erro_PedidoCompraAbertoLista_BotaoEdita

    If lErro = SUCESSO Then
            
        'Se houver alguma data de envio preenchida, chama a tela PedComprasEnv
        If objPedidoCompras.dtDataEnvio <> DATA_NULA Then
        
            Call Chama_Tela("PedComprasEnv", objPedidoCompras)
        
        'Se não houver Data de envio preenchida
        ElseIf objPedidoCompras.dtDataEnvio = DATA_NULA Then
        
            'Se PedCompras for igual a zero, o pedido de compras será avulso
            If objPedidoCompras.iTipoOrig = 0 Then
        
                Call Chama_Tela("PedidoCompras", objPedidoCompras)
        
            'Se o PedCompras for diferente de zero, o pedido de compras será gerado
            Else
            
                Call Chama_Tela("PedComprasGerado", objPedidoCompras)
        
            End If
          
        End If
        
    End If
        
    'Função executada com Sucesso
    PedidoCompraAbertoLista_BotaoEdita = SUCESSO
    
    Exit Function

'--->Aqui se inicia o Tratamento dos Erros
Erro_PedidoCompraAbertoLista_BotaoEdita:

    PedidoCompraAbertoLista_BotaoEdita = gErr

    Select Case gErr
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147434)

    End Select

    Exit Function

End Function

Function AlcadaUsuarioLista_TrataParametros(objTela As Object, colSelecao As Collection) As Long
'Rotina Trata_Parametros especial para o Browser AlcadaUsuarioLista
'Se o browser for chamado sem receber parametros ---> usa o ano atual

Dim lErro As Long
Dim iAno As Integer
Dim sSelecao As String

On Error GoTo Erro_AlcadaUsuarioLista_TrataParametros

    'Se a coleção for nothing ---> atribui nova coleção
    If colSelecao Is Nothing Then Set colSelecao = New Collection
    
    'armazena o ano atual
    iAno = Year(gdtDataAtual)

    'Monta a selecao SQL
    sSelecao = "Ano=" & CStr(iAno)

    'Atribui a nova seleção
    Set objTela.objBrowse1.colSelecao = colSelecao

    If Not (IsMissing(sSelecao)) Then objTela.objBrowse1.sSelecaoSQL2 = sSelecao

    AlcadaUsuarioLista_TrataParametros = SUCESSO

    Exit Function

Erro_AlcadaUsuarioLista_TrataParametros:

    AlcadaUsuarioLista_TrataParametros = gErr

    Select Case gErr

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147435)

    End Select

    Exit Function

End Function

Function ProdutosPCTodos_BotaoEdita(objPedidoCompra As ClassPedidoCompras, lErro As Long) As Long

On Error GoTo Erro_ProdutosPCTodos_BotaoEdita

    If Not (objPedidoCompra Is Nothing) Then
    
        If objPedidoCompra.iStatus = 1 Then gError 123521
    
        If objPedidoCompra.dtDataEnvio <> DATA_NULA And objPedidoCompra.iStatus = 0 Then
            Call Chama_Tela("PedComprasEnv", objPedidoCompra)
        
        Else
            
            If objPedidoCompra.iTipoOrig = 0 Then
                Call Chama_Tela("PedComprasGerado", objPedidoCompra)
            Else
                Call Chama_Tela("PedidoCompras", objPedidoCompra)
            End If

        End If

    Else

        Call Chama_Tela("PedidoCompras")
                        
    End If
    
    ProdutosPCTodos_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_ProdutosPCTodos_BotaoEdita:

    ProdutosPCTodos_BotaoEdita = gErr

    Select Case gErr
    
    Case 123521
        Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO ", gErr, objPedidoCompra.lCodigo)
    
    Case Else
       Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147436)

    End Select

    Exit Function

End Function

Function ItensPedCompraPCLista_BotaoConsulta(objItemPC As ClassItensPedCompraPC, lErro As Long) As Long

Dim objPedidoCompras As New ClassPedidoCompras

On Error GoTo Erro_ItensPedCompraPCLista_BotaoConsulta

    'Se lErro retornar sucesso ----> Executa uma das condições abaixo
    If lErro = SUCESSO Then
    
        objPedidoCompras.lNumIntDoc = objItemPC.lNumIntDoc
    
        Call Chama_Tela("PedComprasCons", objPedidoCompras)
    
    End If
    
    'Função executada com Sucesso
    ItensPedCompraPCLista_BotaoConsulta = SUCESSO
    
    Exit Function

'--->Aqui se inicia o Tratamento dos Erros
Erro_ItensPedCompraPCLista_BotaoConsulta:

    ItensPedCompraPCLista_BotaoConsulta = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147432)

    End Select

    Exit Function

End Function
