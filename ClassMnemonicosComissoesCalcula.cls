VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoComissCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variável utilizada para trabalhar com os controles da tela
'que instanciou essa classe
Private mvarobjTela As Object

'Indica que é para obter o endereço da filialcliente
Const FILIALCLIENTE_OBTEM_ENDERECO = 1

'******************************************************************
'Constantes com os nomes dos campos que precisam ter seus
'valores calculados proporcionalmente para cada produto
Const CAMPO_VALORFRETE As String = "ValorFrete"
Const CAMPO_VALORSEGURO As String = "ValorSeguro"
Const CAMPO_VALORIPI As String = "IPIValor1"
Const CAMPO_VALORICMS As String = "ICMSValor1"
Const CAMPO_VALORDESPESAS As String = "ValorDespesas"
Const CAMPO_VALORPRODUTOS As String = "ValorProdutos"
Const CAMPO_VALORDESCONTO As String = "ValorDesconto"
'******************************************************************

'Variaveis Globais, usadas para armazenar dados de acessos ao BD
Private lUltPedVendaCod As Long
Private iUltFilialEmpPedVendaCod As Integer
Private iUltRegiaoVenda As Integer
Private iUltCondicaoPagto As Integer
Private sUltClienteNomeRed As String
Private iUltFilialClienteCod As Integer
Private sUltProduto As String
Private objUltCliente As ClassCliente
Private objUltFilialCliente As ClassFilialCliente
Private objUltEnderecoFilialCliente As ClassEndereco
Private objUltPedidoVenda As ClassPedidoDeVenda
Private objUltCondicaoPagto As ClassCondicaoPagto
Private objUltProduto As ClassProduto
Private colcolProdutosCategorias As New Collection

'*** INÍCIO PROPERTY GET / PROPERTY SET ***
Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property
'*** FIM PROPERTY GET / PROPERTY SET ***

Public Function Calcula_Mnemonico_Comissoes(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Verifica qual o mnemônico passado e chama a função apropriada para calculá-los
'objMnemonicoValor RECEBE(Input) o mnemônico que será calculado
'e RETORNA(Output) o valor desse mnemônico

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico_Comissoes

    'Faz uma seleção pelo mnemônico
    Select Case objMnemonicoValor.sMnemonico
    
        'Se for RegiaoVenda
        Case MNEMONICO_COMISSOES_REGIAO
            
            'Calcula o valor do mnemonico RegiaoVenda
            lErro = Calcula_MnemonicoComissoes_RegiaoVenda(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98994
        
        'Se for Cliente
        Case MNEMONICO_COMISSOES_CLIENTE
        
            'Calcula o valor do mnemonico Cliente
            lErro = Calcula_MnemonicoComissoes_Cliente(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98995
            
        'Se for FilialCliente
        Case MNEMONICO_COMISSOES_FILIALCLI
        
             'Calcula o valor do mnemonico FilialCliente
             lErro = Calcula_MnemonicoComissoes_FilialCliente(objMnemonicoValor)
             If lErro <> SUCESSO Then gError 98996
        
        'Se for Vendedor
        Case MNEMONICO_COMISSOES_VENDEDOR
        
            'Calcula o valor do mnemonico Vendedor
            lErro = Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98999
            
        'Se for Vendedor
        Case MNEMONICO_COMISSOES_PERCCOMISSCLI
        
            'Calcula o valor do mnemonico Vendedor
            lErro = Calcula_MnemonicoComissoes_PercComissCli(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98999
        
        'Se for ItemCatProduto
        Case MNEMONICO_COMISSOES_ITEMCATPRODUTO
        
            'Calcula o valor do mnemonico ItemCatProduto
            lErro = Calcula_MnemonicoComissoes_ItemCatProduto(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98997
        
        'Se for Produto
        Case MNEMONICO_COMISSOES_PRODUTO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 98998
        
        'Se for Preco de Venda
        Case MNEMONICO_COMISSOES_PRECO_VENDA
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_PrecoVenda(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102017
        
        Case MNEMONICO_COMISSOES_COMISSAO_ITEM
        
            lErro = Calcula_MnemonicoComissoes_ComissaoItem(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102017
        
        Case MNEMONICO_COMISSOES_TABPRECO_ITEM
        
            lErro = Calcula_MnemonicoComissoes_TabPrecoItem(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102017
        
        Case MNEMONICO_COMISSOES_COMISS_TABPRECO_ITEM
        
            lErro = Calcula_MnemonicoComissoes_ComisTabPrecoItem(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102017
        
        Case MNEMONICO_COMISSOES_PRECO_TABELA
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_PrecoTabela(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 126144
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRDESCONTO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORDESCONTO)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_PRODUTO_PERCDESCONTO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_PercDesc(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_VENDEDOR_VINCULO_CODIGO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Vend_Vinculo_Cod(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_VENDEDOR_VINCULO_DESCRICAO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Vend_Vinculo_Desc(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_TABELAPRECO_CODIGO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_TabPreco_Cod(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_TABELAPRECO_DESCRICAO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_TabPreco_Desc(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102076
            
        Case MNEMONICO_COMISSOES_PRODUTO_VLRFRETE
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORFRETE)
            If lErro <> SUCESSO Then gError 102018
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRIPI
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORIPI)
            If lErro <> SUCESSO Then gError 102019
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRICMS
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORICMS)
            If lErro <> SUCESSO Then gError 102019
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLROUTRASDESP
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORDESPESAS)
            If lErro <> SUCESSO Then gError 102020
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRSEGURO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrCampo(objMnemonicoValor, CAMPO_VALORSEGURO)
            If lErro <> SUCESSO Then gError 102021
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRVENDA
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrVenda(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102022
        
        Case MNEMONICO_COMISSOES_PRODUTO_VLRTOTAL
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_Produto_VlrTotal(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102023
        
        Case MNEMONICO_COMISSOES_TIPOPRODUTO
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_TipoProduto(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 131970
        
        Case MNEMONICO_COMISSOES_ITEMCATCLIENTE
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_ItemCatCliente(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 131971
        
        Case MNEMONICO_COMISSOES_CANALVENDA
        
            'Calcula o valor do mnemonico Produto
            lErro = Calcula_MnemonicoComissoes_CanalVenda(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 131971
        
        Case Else
        
            'Retorna uma indicação de que o mnemônico não foi tratado
            Calcula_Mnemonico_Comissoes = MNEMONICOCOMISSOES_NAO_ENCONTRADO
            
            Exit Function
    
    End Select
        
    Calcula_Mnemonico_Comissoes = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico_Comissoes:
    
    Calcula_Mnemonico_Comissoes = gErr
    
    Select Case gErr
    
        Case 98994 To 98999, 102017 To 102023, 102076, 126144
        
        Case 131970, 131971 'Inserido por Wagner
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152192)
    
    End Select
    
End Function

Private Function Calcula_MnemonicoComissoes_RegiaoVenda(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código da região
'objMnemonicoValor RETORNA(Output) o código da região de venda

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Calcula_MnemonicoComissoes_RegiaoVenda

    'Obtém os dados da filial do cliente selecionada
    lErro = Obtem_FilialCliente_Dados(objFilialCliente)
    If lErro <> SUCESSO Then gError 94979
    
    'Guarda na coleção o código da região
    objMnemonicoValor.colValor.Add objFilialCliente.iRegiao
    
    Calcula_MnemonicoComissoes_RegiaoVenda = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_RegiaoVenda:

    Calcula_MnemonicoComissoes_RegiaoVenda = gErr
    
    Select Case gErr
    
        Case 94979
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152193)
    
    End Select
    
    Exit Function

End Function

Public Function Obtem_FilialCliente_Dados(objFilialCliente As ClassFilialCliente, Optional ByVal iObtemEndereco As Integer = 0, Optional objEndereco As ClassEndereco) As Long
'Devolve os dados da filialcliente que está selecionada na tela
'objFilialCliente RETORNA(Output) os dados da filialcliente
'iObtemEndereco RECEBE(Input) uma flag indicando se é para ler os dados de endereco
'objEndereco RETORNA(Output) os dados do endereço da filialcliente, caso este tenha sido lido

Dim lErro As Long
Dim sClienteNomeRed As String
Dim bLeEnderecoBD As Boolean

On Error GoTo Erro_Obtem_FilialCliente_Dados

    'Guarda o nome reduzido do cliente que está na tela
    sClienteNomeRed = objTela.Controls("Cliente")
    
    'Guarda no obj o código da filial que está na tela
    objFilialCliente.iCodFilial = Codigo_Extrai(objTela.Controls("Filial"))
    
    'Se o cliente e a filial estão preenchidos
    If Len(Trim(sClienteNomeRed)) > 0 And (objFilialCliente.iCodFilial > 0) Then
    
        'Se mudou a filial ou o cliente
        If (objFilialCliente.iCodFilial <> iUltFilialClienteCod) Or (sClienteNomeRed <> sUltClienteNomeRed) Then
        
            'Se mudou o cliente
            If sClienteNomeRed <> sUltClienteNomeRed Then
            
                'Limpa os dados armazenados do último cliente lido
                Call Limpa_Dados_UltCliente
                
            End If
        
            '*** LEITURA DOS DADOS DA FILIAL ***
            
            'Lê os dados da filial do cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sClienteNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 94983
            
            'Se não encontrou a filial => erro
            If lErro = 17660 Then gError 94984
            '*** FIM DA LEITURA DA FILIAL ***
            
            'indica que é para ler o endereço a partir do BD (só será utilizado,
            'se a função for executar a leitura do BD)
            bLeEnderecoBD = True
            
            'Guarda na variável o código da última filialcliente que teve seus dados
            'lidos no BD
            iUltFilialClienteCod = objFilialCliente.iCodFilial
            
            'Instancia objUltFilialCliente com os dados da última filialcliente lida
            Set objUltFilialCliente = objFilialCliente
        
        'Senão, ou seja, se tanto a filial quanto o cliente, permanecem os mesmos
        Else
        
            'Instancia objFilialCliente com os dados da última filial lida
            Set objFilialCliente = objUltFilialCliente
        
        End If
        
        'teste de seguranca.... por tulio em 8/7
        If Not objUltEnderecoFilialCliente Is Nothing Then
            bLeEnderecoBD = False
        Else
            bLeEnderecoBD = True
        End If
        
        'Obtém o endereço da filial
        lErro = Obtem_FilialCliente_Endereco(iObtemEndereco, objFilialCliente.lEndereco, objEndereco, bLeEnderecoBD)
        If lErro <> SUCESSO Then gError 94989
    
    'Senão, ou seja, se o cliente ou a filial não estão preenchidos
    Else
        
        'Se o cliente não está preenchido
        If Len(Trim(sClienteNomeRed)) = 0 Then
        
            'Limpa os dados armazenados do último cliente lido
            Call Limpa_Dados_UltCliente
            
        End If
        
        'Limpa a variável que guarda o código da última filialcliente que teve
        'seus dados lidos no BD
        iUltFilialClienteCod = 0
        
        'Limpa o obj que guarda os dados do último cliente lido
        Set objUltFilialCliente = Nothing
        
        'Limpa o obj que guarda os dados do último endereço de filialcliente lido
        Set objUltEnderecoFilialCliente = Nothing
    
    End If
    
    Obtem_FilialCliente_Dados = SUCESSO
    
    Exit Function

Erro_Obtem_FilialCliente_Dados:

    Obtem_FilialCliente_Dados = gErr
    
    Select Case gErr
    
        Case 94980, 94983, 94989
        
        Case 94984
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE1", gErr, objFilialCliente.iCodFilial, sClienteNomeRed)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152194)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Limpa_Dados_UltCliente()

    'Limpa a variável que guarda o nome reduzido do último cliente que teve
    'seus dados lidos no BD
    sUltClienteNomeRed = STRING_VAZIO
    
    'Limpa o obj que guarda os dados do último cliente lido
    Set objUltCliente = Nothing

End Sub

Private Sub Limpa_Dados_UltCondicaoPagto()

    'Limpa a variável que guarda o a condicao de pagamento da ultima cond lida
    iUltCondicaoPagto = 0
    
    'Limpa o obj que guarda os dados da ultima cond. lida...
    Set objUltCondicaoPagto = Nothing

End Sub

Private Function Calcula_MnemonicoComissoes_Cliente(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados do cliente que está na tela e devolve o código
'objMnemonicoValor RETORNA(Output) o código da região de venda

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Calcula_MnemonicoComissoes_Cliente

    'Obtém os dados do cliente selecionado
    lErro = Cliente_Obtem_Dados(objCliente)
    If lErro <> SUCESSO Then gError 94981
    
    'Guarda na coleção o código do cliente
    objMnemonicoValor.colValor.Add objCliente.lCodigo
    
    Calcula_MnemonicoComissoes_Cliente = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Cliente:

    Calcula_MnemonicoComissoes_Cliente = gErr
    
    Select Case gErr
    
        Case 94981

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152195)
    
    End Select
    
    Exit Function

End Function

Public Function Cliente_Obtem_Dados(objCliente As ClassCliente) As Long
'Devolve os dados do cliente que está selecionado na tela
'objCliente RETORNA(Output) os dados do cliente

Dim lErro As Long
Dim sClienteNomeRed As String

On Error GoTo Erro_Cliente_Obtem_Dados

    'Guarda o nome reduzido do cliente que está na tela
    sClienteNomeRed = objTela.Controls("Cliente")
    
    'Se o cliente foi informado
    If Len(Trim(sClienteNomeRed)) > 0 Then
    
        'Se o nome reduzido do cliente selecionado não se alterou
        If sClienteNomeRed = sUltClienteNomeRed Then
        
            'Significa que o cliente não foi alterado e o código é o mesmo calculado anteriormente
            Set objCliente = objUltCliente
        
        'Senão, significa que o cliente foi alterado e é necessário ler o cliente novamente
        Else
        
            'Guarda o nome reduzido do cliente no obj
            objCliente.sNomeReduzido = sClienteNomeRed
            
            'Lê os dados do cliente a partir do nome reduzido guardado no obj
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 94977
            
            'Se não encontrou o cliente
            If lErro = 12348 Then gError 94978
            
            'Guarda na variável o nome reduzido do último cliente que teve seus dados
            'lidos no BD
            sUltClienteNomeRed = sClienteNomeRed
            
            'Instancia objUltCliente com os dados do último cliente lido
            Set objUltCliente = objCliente
        
        End If
    
    'Senão, ou seja, se o cliente não foi informado
    Else
    
        'Limpa os dados armazenados do último cliente lido
        Call Limpa_Dados_UltCliente
        
    End If

    Cliente_Obtem_Dados = SUCESSO
    
    Exit Function

Erro_Cliente_Obtem_Dados:

    Cliente_Obtem_Dados = gErr
    
    Select Case gErr
    
        Case 94977
        
        Case 94978
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, sClienteNomeRed)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152196)
    
    End Select
    
    Exit Function
    
End Function

'criada por tulio em 8/7
Public Function CondPagto_Obtem_Dados(objCondicaoPagto As ClassCondicaoPagto, Optional iCondPagto As Integer = 0) As Long
'Devolve os dados da condpagto que está selecionado na tela
'objCondpagto eh parametro de output que volta os dados da condicao de pagamento
'iCondPagto eh parametro opcional utilizado pela tela NFiscalPedido, por tulio em 01/08/02

Dim lErro As Long

On Error GoTo Erro_CondPagto_Obtem_Dados

    If iCondPagto = 0 Then
    
        'Guarda a condicao de pagamento que esta na tela... caso a mesma exista...
        objCondicaoPagto.iCodigo = Codigo_Extrai((objTela.Controls("CondicaoPagamento").Text))
    
    Else
            
        objCondicaoPagto.iCodigo = iCondPagto
    
    End If
    
    'se a condicao foi informada
    If objCondicaoPagto.iCodigo > 0 Then
    
        'se o cod. da condicao nao se alterou
        If objCondicaoPagto.iCodigo = iUltCondicaoPagto Then
        
            'faz com que aponte para o ultimo lido, evitando assim, uma nova leitura
            Set objCondicaoPagto = objUltCondicaoPagto
            
        'significa que a condicao foi alterada e que uma nova leitura sera necessaria..
        Else
            
            'le os dados da condicao a partir do codigo...
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 101761
        
            'Se a condição de pagamento não está cadastrada, erro
            If lErro = 19205 Then gError 101762
            
            'Guarda na variável o codigo da ultima cond. lida...
            iUltCondicaoPagto = objCondicaoPagto.iCodigo
            
            'Instancia objUltCondicaoPagto com os dados da ultima Condicao lida...
            Set objUltCondicaoPagto = objCondicaoPagto
        
        End If
    
    'Senao, se a condicao nao foi informada..
    Else
    
        'Limpa os dados remanescentes da ultima condicao lida
        Call Limpa_Dados_UltCondicaoPagto
        
    End If

    CondPagto_Obtem_Dados = SUCESSO
    
    Exit Function

Erro_CondPagto_Obtem_Dados:

    CondPagto_Obtem_Dados = gErr
    
    Select Case gErr
    
        Case 101761
        
        Case 101762
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152197)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_FilialCliente(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim iFilialClienteCod As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_FilialCliente

    'Se a filial do cliente está preenchida
    If Len(Trim(objTela.Controls("Filial"))) > 0 Then
        
        'Guarda na coleção o código da filial
        objMnemonicoValor.colValor.Add Codigo_Extrai(objTela.Controls("Filial"))
        
        'Se a filial do cliente mudou
        If iUltFilialClienteCod <> Codigo_Extrai(objTela.Controls("Filial")) Then
        
            'Limpa os dados armazenados da última filialcliente lida
            Call Limpa_Dados_ultFilialCliente
            
            'Guarda a nova filial do cliente
            iUltFilialClienteCod = Codigo_Extrai(objTela.Controls("Filial"))
        
        End If
            
    'Senão
    Else
            
        'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
        objMnemonicoValor.colValor.Add 0
        
        'Limpa os dados armazenados da última filialcliente lida
        Call Limpa_Dados_ultFilialCliente
    
    End If
    
    Calcula_MnemonicoComissoes_FilialCliente = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_FilialCliente:

    Calcula_MnemonicoComissoes_FilialCliente = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152198)
    
    End Select
    
    Exit Function

End Function

Private Sub Limpa_Dados_ultFilialCliente()

    'Limpa a variável que guarda o código da última filialcliente que teve
    'seus dados lidos no BD
    iUltFilialClienteCod = 0
    
    'Limpa o obj que guarda os dados do último cliente lido
    Set objUltFilialCliente = Nothing

End Sub

Private Function Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Calcula_MnemonicoComissoes_Vendedor

    'Alterada por Luiz Nogueira em 22/04/04 para atender às comissões do loja
    'Se estiver calculando comissões do loja, o código do vendedor está na própria tela
    If objTela.Caption = NOME_TELA_COMISSOESCALCULALOJA Then
    
        objMnemonicoValor.colValor.Add StrParaInt(objTela.Controls("VendedorCupom").Text)
    
    
    Else
        
        'Obtém os dados da filial do cliente selecionada
        lErro = Obtem_FilialCliente_Dados(objFilialCliente)
        If lErro <> SUCESSO Then gError 94985
    
        'Guarda na coleção o código do vendedor
        objMnemonicoValor.colValor.Add objFilialCliente.iVendedor
    
    End If
    
    Calcula_MnemonicoComissoes_Vendedor = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_Vendedor:

    Calcula_MnemonicoComissoes_Vendedor = gErr
    
    Select Case gErr
    
        Case 94985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_PercComissCli(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o percentual de comissão
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Calcula_MnemonicoComissoes_PercComissCli
        
    'Obtém os dados da filial do cliente selecionada
    lErro = Obtem_FilialCliente_Dados(objFilialCliente)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Guarda na coleção o código do vendedor
    objMnemonicoValor.colValor.Add objFilialCliente.dComissaoVendas
    
    Calcula_MnemonicoComissoes_PercComissCli = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_PercComissCli:

    Calcula_MnemonicoComissoes_PercComissCli = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_ItemCatProduto(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Obtém o valor do item da categoria passada como parâmetro para cada produto do grid de itens
'objMnemonicoValor RETORNA(Output) o item da categoria para cada produto no grid

Dim bAchou As Boolean
Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iIndice As Integer
Dim iLinha As Integer
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim colProdutoCategoria As Collection
Dim iProdPreenchido As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_ItemCatProduto

    'Para cada linha do grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Se o produto está preenchido
        If Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0 Then
        
            'Guarda o produto informado na linha atual
            sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
            
            'formata o produto para poder manipula-lo em memoria
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdPreenchido)
            If lErro <> SUCESSO Then gError 101561
            
            'Procura o produto passado como parâmetro na coleção de Categorias de Produtos
            For iIndice = 1 To colcolProdutosCategorias.Count
            
                'Instancia o obj com o primeiro item da coleção de categorias atual
                Set objProdutoCategoria = colcolProdutosCategorias.Item(iIndice).Item(1)
                
                'Se o produto é o mesmo da linha atual
                If objProdutoCategoria.sProduto = sProdutoFormatado Then
                
                    'Procura a categoria na coleção de categorias do produto
                    For Each objProdutoCategoria In colcolProdutosCategorias.Item(iIndice)
                    
                        'Se a categoria atual é igual à categoria passada como parâmetro
                        If (UCase(objProdutoCategoria.sCategoria) = UCase(objMnemonicoValor.vParam(1))) Then
                        
                            'Guarda na coleção o item da categoria para o par produto, categoria passado
                            objMnemonicoValor.colValor.Add objProdutoCategoria.sItem
                            
                            'Indica que já encontrou o item da categoria
                            bAchou = True
                            
                            'Sai do loop que procura a categoria na coleção de categorias do produto
                            Exit For
                        
                        End If
                    
                    Next
                    
                    'Se já achou o item => sai do loop que procura o produto na coleção de categorias de produtos
                    If bAchou Then Exit For
                
                End If
            
            Next
            
            'Se ainda não achou o item de categoria
            If Not bAchou Then
            
                'Instancia uma nova coleção de categorias de produto
                Set colProdutoCategoria = New Collection
                
                'Lê todas as categorias para o produto passado como parâmetro
                lErro = CF("ProdutoCategorias_Le_Produto", sProdutoFormatado, colProdutoCategoria)
                If lErro <> SUCESSO Then gError 94986
                
                'Procura a categoria passada como parâmetro na coleção de categorias lidas no BD
                For Each objProdutoCategoria In colProdutoCategoria
                
                    'Se a categoria atual é igual à categoria passada como parâmetro
                    If (objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)) Then
                    
                        'Guarda na coleção o item da categoria para o par produto, categoria passado
                        objMnemonicoValor.colValor.Add objProdutoCategoria.sItem
                        
                        'Indica que encontrou o item da categoria
                        bAchou = True
                        
                        'Guarda a coleção na coleção de categorias de produtos
                        colcolProdutosCategorias.Add colProdutoCategoria
                        
                        'Sai do loop que procura a categoria na coleção de categorias do produto
                        Exit For
                    
                    End If
                
                Next
                
                'Se não achou a categoria => Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
                If Not bAchou Then objMnemonicoValor.colValor.Add STRING_VAZIO
            
            End If
        
        'Senão, ou seja, se o produto não está preenchido
        Else
        
            'Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add STRING_VAZIO
        
        End If
        
        'Limpa a variável que indica se o item de categoria foi encontrado
        bAchou = False
            
    Next

    Calcula_MnemonicoComissoes_ItemCatProduto = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_ItemCatProduto:

    Calcula_MnemonicoComissoes_ItemCatProduto = gErr
    
    Select Case gErr
    
        Case 94986, 101561
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152200)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_Produto(objMnemonicoValor As ClassMnemonicoValor) As Long
'Adiciona o código dos produtos do grid à coleção de mnemônicos
'objMnemonicoValor RETORNA(Output) os códigos do produtos do grid de itens

Dim iLinha As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_Produto

    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        'Se o produto está preenchido
        If Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0 Then
            
            'Adiciona o código do produto à coleção
            objMnemonicoValor.colValor.Add objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'Senão
        Else
            
            'Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add STRING_VAZIO
        End If
        
    Next

    Calcula_MnemonicoComissoes_Produto = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Produto:
    
    Calcula_MnemonicoComissoes_Produto = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152201)
    
    End Select
    
    Exit Function
    
End Function

'Alterada por Luiz em 30/05/03
Public Function Calcula_MnemonicoComissoes_PrecoVenda(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o preço de venda para cada produto no grid de itens
'objMnemonicoValor RETORNA(Output) os preços de vendas encontrados

Dim iLinha As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUM As String
Dim dFator As Double
Dim dPrecoVenda As Double
Dim lErro As Long

On Error GoTo Erro_Calcula_MnemonicoComissoes_PrecoVenda
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        'Guarda o produto da linha atual
        sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'formata o produto para que o mesmo possa ser usado como filtro
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 102240
        
        'Se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            'Guarda o código do produto no obj
            objProduto.sCodigo = sProdutoFormatado
            
            'Obtém os dados do produto
            lErro = Produto_Obtem_Dados(objProduto)
            If lErro <> SUCESSO Then gError 102241
            
            'Guarda a U.M. do produto da linha atual
            sUM = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_UnidadeMed_Col)
            
            'Encontra o fator de conversão da UM que está no grid para a UM base de venda
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 102242

            'guarda o preco unitario q eh parametro de input e output da funcao q sera chamada na prox. linha
            dPrecoVenda = StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_ValorUnitario_Col)) / dFator

            'Adiciona o preço unitário do produto à coleção
            objMnemonicoValor.colValor.Add dPrecoVenda
        
        'Senão
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
        End If
        
    Next

    Calcula_MnemonicoComissoes_PrecoVenda = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_PrecoVenda:

    Calcula_MnemonicoComissoes_PrecoVenda = gErr
    
    Select Case gErr
    
        Case 102240 To 102242
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152202)
    
    End Select
    
    Exit Function
    
End Function

Private Function Obtem_FilialCliente_Endereco(ByVal iObtemEndereco As Integer, ByVal lEndereco As Long, objEndereco As ClassEndereco, ByVal bLeEnderecoBD As Boolean) As Long
'Verifica se deve ser retornado um endereço e em caso positivo, preenche objEndereco
'iObtemEndereco RECEBE(Input) uma flag indicando se o endereço deve ser obtido
'lEndereco RECEBE(Input) o código do endereço que será lido
'objEndereco RETORNA(Output) os dados do endereço lido
'bLeEnderecoBD RECEBE(Input) uma flag que indica se o endereço deve ser lido do BD

Dim lErro As Long

On Error GoTo Erro_Obtem_FilialCliente_Endereco

    'Se é para obter o endereço da filial
    If (iObtemEndereco = FILIALCLIENTE_OBTEM_ENDERECO) Then
            
        'Se o endereco deve ser lido do bd
        If (bLeEnderecoBD) Then
        
            'Guarda no objEndereco o código do endereço que deve ser lido
            objEndereco.lCodigo = lEndereco
            
            'Lê o endereço da filialcliente
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 94987
            
            'Se não encontrou o endereço no BD => Erro
            If lErro = 12309 Then gError 94988
            
            'Instancia o obj que guarda os dados do último endereço de filialcliente lido,
            'com os dados do endereço atual
            Set objUltEnderecoFilialCliente = objEndereco
        
        'Senão, ou seja, se não precisa ler o endereço do BD
        Else
        
            'Instancia o objEndereco com os dados do último endereço lido
            Set objEndereco = objUltEnderecoFilialCliente
            
        End If
    
    End If
        
    Obtem_FilialCliente_Endereco = SUCESSO
    
    Exit Function
    
Erro_Obtem_FilialCliente_Endereco:

    Obtem_FilialCliente_Endereco = gErr
    
    Select Case gErr
    
        Case 94987
        
        Case 94988
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152203)
    
    End Select

    Exit Function
    
End Function

Public Function Obtem_PedidoVenda_Dados(objPedidoVenda As ClassPedidoDeVenda) As Long
'Devolve os dados do pedidovenda que está selecionado na tela
'objPedidoVenda RETORNA(Output) os dados do pedido de venda lido

Dim lErro As Long

On Error GoTo Erro_Obtem_PedidoVenda_Dados

    'Guarda no obj o código do pedido que está selecionado na tela
    objPedidoVenda.lCodigo = objTela.glCodPedido
    
    'Guarda no obj o código da filialpedido que está selecionada na tela
    objPedidoVenda.iFilialEmpresa = objTela.giFilialPedido
    
    'Se o código do pedido e da filialpedido estão preenchidos
    If (objPedidoVenda.lCodigo > 0) And (objPedidoVenda.iFilialEmpresa > 0) Then
    
        'Se mudou o código do pedido de venda
        If (objPedidoVenda.lCodigo <> lUltPedVendaCod) Or (objPedidoVenda.iFilialEmpresa <> iUltFilialEmpPedVendaCod) Then
            
            'Significa que os dados da última leitura não servem mais e o pedido deve
            'ser lido novamente
            lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
            If lErro <> SUCESSO And lErro <> 26509 Then gError 102003
            
            'Se não encontrou o pedido => erro
            If lErro = 26509 Then gError 102004
            
            'Guarda na variável o código do último pedido que teve seus dados
            'lidos no BD
            lUltPedVendaCod = objPedidoVenda.lCodigo
            
            'Guarda na variável o código da última filialpedido que teve seus dados
            'lidos no BD
            iUltFilialEmpPedVendaCod = objPedidoVenda.iFilialEmpresa
            
            'Instancia objUltPedidoVenda com os dados do último pedido de venda lido
            Set objUltPedidoVenda = objPedidoVenda
        
        'Senão, ou seja, se o código do pedido e da filialpedido são os mesmos
        Else
            
            'Instancia objPedidoVenda com os dados do último pedido de venda lido
            Set objPedidoVenda = objUltPedidoVenda
        
        End If
    
    'Senão, ou seja se código do pedido ou da filial estão em branco
    Else
    
        'Limpa a variável que guarda o código do último pedido de venda que teve
        'seus dados lidos no BD
        lUltPedVendaCod = 0
        
        'Limpa a variável que guarda o código da última filial de pedido de venda que teve
        'seus dados lidos no BD
        iUltFilialEmpPedVendaCod = 0
        
        'Limpa o obj que guarda os dados do último pedido venda lida
        Set objUltPedidoVenda = Nothing
    
    End If
        
    Obtem_PedidoVenda_Dados = SUCESSO
    
    Exit Function
    
Erro_Obtem_PedidoVenda_Dados:

    Obtem_PedidoVenda_Dados = gErr
    
    Select Case gErr
    
        Case 102003
        
        Case 102004
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152204)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_Produto_VlrCampo(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sCampoValor As String) As Long
'Devolve o valor do campo passado em sCampo calculado proporcionalmente para cada produto do grid itens
'objMnemonicoValor RETORNA(Output) os valores encontrados

Dim iLinha As Integer
Dim lErro As Long
Dim dFator As Double
Dim dValorTotProdutos As Double
Dim dValorProduto As Double
Dim dValorCampo As Double
Dim dValorCampoProduto As Double
Dim iIndice As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_Produto_VlrCampo

    'Obtém o valor total dos produtos que será utilizado para calcular a proporção
    'Para cada produto no grid de itens
    For iIndice = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Acumula o valor total dos produtos
        dValorTotProdutos = dValorTotProdutos + StrParaDbl(objTela.Controls("GridItens").TextMatrix(iIndice, objTela.iGrid_PrecoTotal_Col))
    
    Next
    
    Select Case sCampoValor

        Case CAMPO_VALORFRETE, CAMPO_VALORSEGURO, CAMPO_VALORDESPESAS, CAMPO_VALORDESCONTO

            'Guarda o valor total do campo em questão
            dValorCampo = StrParaDbl(objTela.Controls(sCampoValor).Text)
    
        Case CAMPO_VALORIPI, CAMPO_VALORICMS
            
            'Guarda o valor total do campo em questão
            dValorCampo = StrParaDbl(objTela.Controls(sCampoValor).Caption)
    
    End Select
    
    'Para cada linha do grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        'Só faz sentido calcular o valor proporcional do campo em questão se=>
        'O produto estiver preenchido
        'O valor do campo em questão for diferente de zero
        'O valor total de produtos for diferente de zero
        If (Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0) And (dValorTotProdutos <> 0) And (dValorCampo <> 0) Then
        
            'Guarda o valor total desse produto
            dValorProduto = StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PrecoTotal_Col))
            
            'Encontra a proporção do frete que será apropriada para esse produto
            dFator = dValorProduto / dValorTotProdutos
            
            'Calcula o valor do campo em questão que será apropriado para o produto
            dValorCampoProduto = dValorCampo * dFator
            
            'Guarda o valor calculado para o campo em questão
            objMnemonicoValor.colValor.Add dValorCampoProduto
        
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0

        End If
        
    Next
    
    Calcula_MnemonicoComissoes_Produto_VlrCampo = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Produto_VlrCampo:

    Calcula_MnemonicoComissoes_Produto_VlrCampo = gErr
    
    Select Case gErr
    
        Case 102075
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152205)
    
    End Select
    
    Exit Function

End Function

Private Function Calcula_MnemonicoComissoes_Produto_VlrVenda(objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o valor total de venda para cada produto no grid de itens
'(Qunatidade * Preço Unitário, ou seja, o preço total sem incluir a proporção de frete, seguro, etc.)
'objMnemonicoValor RETORNA(Output) os valores encontrados

Dim lErro As Long, iLinha As Integer
Dim bFat As Boolean

On Error GoTo Erro_Calcula_MnemonicoComissoes_Produto_VlrVenda

    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Se o produto está preenchido
        If Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0 Then
    
            bFat = True
            If Existe_objTribTab_Tela Then
        
                If objTela.gobjTribTab.iTipoTela = objTela.gobjTribTab.TIPOTELA_NF Then
                
                    If objTela.gobjTribTab.gcolDocTribItem.Count >= iLinha Then
                    
                        lErro = objTela.gobjTribTab.Testa_Item_Fat(iLinha, bFat)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                    End If
            
                End If
            
            End If
            
            If bFat Then
            
                'Adiciona à coleção o preço de venda do produto
                '(Quantidade * Preço Unitário, ou seja, o preço total sem incluir a proporção de frete, seguro, etc.)
                objMnemonicoValor.colValor.Add StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PrecoTotal_Col))
        
            Else
            
                'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
            
        End If
    
    Next
    
    Calcula_MnemonicoComissoes_Produto_VlrVenda = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Produto_VlrVenda:

    Calcula_MnemonicoComissoes_Produto_VlrVenda = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152206)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_Produto_VlrTotal(objMnemonicoValor As ClassMnemonicoValor) As Long
'Calcula o valor total de cada produto no grid de itens incorporando a proporção de desconto, frete, despesas, seguro e ipi
'objMnemonicoValor RETORNA(Output) os valores calculados

Dim iLinha As Integer
Dim dValorProduto As Double
Dim lErro As Long
Dim bFat As Boolean

On Error GoTo Erro_Calcula_MnemonicoComissoes_Produto_VlrTotal
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Se o produto está preenchido
        If Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0 Then
    
            bFat = True
            If Existe_objTribTab_Tela Then
        
                If objTela.gobjTribTab.iTipoTela = objTela.gobjTribTab.TIPOTELA_NF Then
                
                    If objTela.gobjTribTab.gcolDocTribItem.Count >= iLinha Then
                    
                        lErro = objTela.gobjTribTab.Testa_Item_Fat(iLinha, bFat)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                    End If
            
                End If
            
            End If
            
            If bFat Then
            
                'Guarda o valor total desse produto
                dValorProduto = StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PrecoTotal_Col))
                
                'Apropria ao valor do produto as devidas proporções de desconto, frete, despesa e ipi
                lErro = Produto_VlrTotal_Calcula(dValorProduto)
                If lErro <> SUCESSO Then gError 102074
                
                'Adiciona à coleção o valor total do produto
                '(Quantidade * Preço Unitário + a proporção de frete, seguro, ipi e outras despesas)
                objMnemonicoValor.colValor.Add dValorProduto
                
            Else
            
                'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
                objMnemonicoValor.colValor.Add 0
            
            End If
        
        'Senão
        Else
        
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
        
        End If
    
    Next
    
    Calcula_MnemonicoComissoes_Produto_VlrTotal = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Produto_VlrTotal:

    Calcula_MnemonicoComissoes_Produto_VlrTotal = gErr
    
    Select Case gErr
    
        Case 102074, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152207)
    
    End Select
    
    Exit Function
    
End Function

'##############################################################
'Inserido por Wagner
Private Function Calcula_MnemonicoComissoes_TipoProduto(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long

Dim iLinha As Integer
Dim sProdutoTela As String
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_Calcula_MnemonicoComissoes_TipoProduto
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Se o produto está preenchido
        If Len(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)) > 0 Then
    
            sProdutoTela = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
            
            lErro = CF("Produto_Formata", sProdutoTela, sProdutoFormatado, iProdPreenchido)
            If lErro <> SUCESSO Then gError 131971
            
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = Produto_Obtem_Dados(objProduto)
            If lErro <> SUCESSO Then gError 131972
            
            'Adiciona à coleção o tipo do produto (Código do Tipo)
            objMnemonicoValor.colValor.Add objProduto.iTipo
        
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
            
        End If
    
    Next
    
    Calcula_MnemonicoComissoes_TipoProduto = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_TipoProduto:

    Calcula_MnemonicoComissoes_TipoProduto = gErr
    
    Select Case gErr
    
        Case 131971, 131972
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152208)
    
    End Select
    
    Exit Function

End Function

Private Function Calcula_MnemonicoComissoes_ItemCatCliente(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long

Dim bAchou As Boolean
Dim lErro As Long
Dim objClienteCategoria As New ClassFilialCliCategoria
Dim colClienteCategoria As New Collection
Dim iProdPreenchido As Integer
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Calcula_MnemonicoComissoes_ItemCatCliente
    
    lErro = Obtem_FilialCliente_Dados(objFilialCliente)
    If lErro <> SUCESSO Then gError 131974
               
    If objFilialCliente.lCodCliente <> 0 And objFilialCliente.iCodFilial <> 0 Then
                
        'Lê todas as categorias para a filial cliente passada como parâmetro
        lErro = CF("ClienteCategorias_Le_FilialCliente", objFilialCliente, colClienteCategoria)
        If lErro <> SUCESSO Then gError 131975
        
        bAchou = False
        
        'Procura a categoria passada como parâmetro na coleção de categorias lidas no BD
        For Each objClienteCategoria In colClienteCategoria
        
            'Se a categoria atual é igual à categoria passada como parâmetro
            If (objClienteCategoria.sCategoria = objMnemonicoValor.vParam(1)) Then
            
                'Guarda na coleção o item da categoria para o par FilialCliente, Categoria passado
                objMnemonicoValor.colValor.Add objClienteCategoria.sItem
                
                'Indica que encontrou o item da categoria
                bAchou = True
                
                'Sai do loop que procura a categoria na coleção de categorias do produto
                Exit For
            
            End If
        
        Next
        
        'Se não achou a categoria => Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
        If Not bAchou Then objMnemonicoValor.colValor.Add STRING_VAZIO
        
    'Senão, ou seja, se o Cliente não está preenchido
    Else
    
        'Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
        objMnemonicoValor.colValor.Add STRING_VAZIO
    
    End If

    Calcula_MnemonicoComissoes_ItemCatCliente = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_ItemCatCliente:

    Calcula_MnemonicoComissoes_ItemCatCliente = gErr
    
    Select Case gErr
    
        Case 131974, 131975
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152209)
    
    End Select
    
    Exit Function
    
End Function
'##############################################################


Private Function Produto_VlrTotal_Calcula(dValor As Double) As Long
'Apropria o valor de frete, ipi, despesas e seguro ao valor total de venda passado como parâmetro
'dValorVenda RECEBE(Input) o valor ao qual serão incorporados frete, ipi, despesas e seguro
'e RETORNA(Output) o valor total

Dim dFator As Double
Dim dValorTotProdutos As Double
Dim adValorCampo(1 To 5) As Double
Dim iLinha As Integer
Dim iIndice As Integer

On Error GoTo Erro_Produto_VlrTotal_Calcula

    'Se o valor passo é zero => não há nada a ser apropriado a ele
    If dValor = 0 Then Exit Function
    
    'Para cada produto no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Acumula o valor total dos produtos
        dValorTotProdutos = dValorTotProdutos + StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PrecoTotal_Col))
    
    Next
    
    'Guarda o valor total de frete, ipi, despesas, seguro e desconto
    adValorCampo(1) = StrParaDbl(objTela.Controls(CAMPO_VALORFRETE).Text)
    adValorCampo(2) = StrParaDbl(objTela.Controls(CAMPO_VALORIPI).Caption)
    adValorCampo(3) = StrParaDbl(objTela.Controls(CAMPO_VALORDESPESAS).Text)
    adValorCampo(4) = StrParaDbl(objTela.Controls(CAMPO_VALORSEGURO).Text)
    adValorCampo(5) = -StrParaDbl(objTela.Controls(CAMPO_VALORDESCONTO).Text)

    'Encontra o fator que indica a proporção do produto em relação ao total de produtos
    dFator = dValor / dValorTotProdutos
            
    'Para cada campo que deve ter seu valor calculado proporcionalmente
    For iIndice = LBound(adValorCampo) To UBound(adValorCampo)
    
        'Incorpora ao valor passado, os valores de cada campo
        dValor = dValor + (adValorCampo(iIndice) * dFator)
    
    Next

    Produto_VlrTotal_Calcula = SUCESSO
    
    Exit Function
    
Erro_Produto_VlrTotal_Calcula:

    Produto_VlrTotal_Calcula = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152210)
    
    End Select

End Function

Private Sub Class_Terminate()

On Error GoTo Erro_Class_Terminate

    'Libera a memória que estava alocada para os objetos e coleções globais à classe
    Set objUltCliente = Nothing
    Set objUltFilialCliente = Nothing
    Set colcolProdutosCategorias = Nothing
    
    Exit Sub

Erro_Class_Terminate:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152211)
    
    End Select
    
    Exit Sub

End Sub

'Tulio aproveitou a funcao da tela de pedido venda da inpal 10/07
'Foi colocada aki pois pode surgir necessidade de utilizacao por outros clientes ou por modulos da versao2...
'6 - Marcio - 08/2000 - incluido tratamento de taxa financeira específica/padrão de cliente
Sub Calcula_TaxaFinanceira(dDias As Double, dTaxaFin As Double, dTaxaFinDia As Double)
'Calcula taxa financeira diária a partir dos número de dias e da taxa financeira mensal passada
        
Dim dTaxaUmDia As Double

    'Calcula os juros compostos ao dia a partir da taxa financeira (ao mês) passada
    'TaxaFin ao dia = (1 + TaxaFin ao mês) elevado a (1 / 30 dias) -1
    dTaxaUmDia = ((1 + dTaxaFin) ^ (0.0333)) - 1
    
    'Calcula a taxa financeira para iDias a partir da taxa financeira ao dia
    dTaxaFinDia = ((1 + dTaxaUmDia) ^ dDias) - 1
            
End Sub

'Tulio aproveitou a funcao da tela de pedido venda da inpal 10/07
'Foi colocada aki pois pode surgir necessidade de utilizacao por outros clientes ou por modulos da versao2...
'6 - Marcio - 08/2000 - incluido tratamento de taxa financeira específica/padrão de cliente
Sub CalculaDias_CondPagamento(objCondPagto As ClassCondicaoPagto, dDias As Double)
    
Dim iDiaParcelaAtual As Integer
Dim iTotalDias As Integer
Dim iNumeroParcelas As Integer
Dim iPeso As Integer
Dim iIntervalo As Integer

    'Se a condição de pagamento for à vista
    If objCondPagto.iCodigo = COD_A_VISTA Or objCondPagto.iCodigo = 0 Then
        
        dDias = 0
    
    'Se não for à vista
    Else
        
        'Se a condição de pagamento for Mensal
        If objCondPagto.iMensal = 1 Then
            iIntervalo = 30
        Else
            iIntervalo = objCondPagto.iIntervaloParcelas
        End If
        
        'Guarda o número de parcelas tirando a primeira
        iNumeroParcelas = objCondPagto.iNumeroParcelas - 1
                    
        'Se possui mais de duas parcelas, calcula a média
        If iNumeroParcelas > 1 Then
        
            'Calcula total de dias da condição de pagamento
            Do While iNumeroParcelas >= 0
                
                'Calcula o número de dias que faltam para chegar a parcela em questão
                iDiaParcelaAtual = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo * iNumeroParcelas)
                
                'Acumula o número de dias de todas as parcelas
                iTotalDias = iTotalDias + iDiaParcelaAtual
                
                'Decrementa o número de parcelas
                iNumeroParcelas = iNumeroParcelas - 1
                                        
            Loop
            
            'Calcula a média ponderada de dias
            dDias = iTotalDias / objCondPagto.iNumeroParcelas
        
        'Se possui exatamente duas parcelas
        ElseIf iNumeroParcelas = 1 Then
            
            dDias = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo / 2)
        
        'Se só possui exatamente uma parcela
        Else
            
            dDias = objCondPagto.iDiasParaPrimeiraParcela
            
        End If
        
    End If
    
End Sub

'Incluida por Luiz em 30/05/03
Public Function Produto_Obtem_Dados(objProduto As ClassProduto) As Long
'Devolve os dados do produto passado como parâmetro
'objProduto RECEBE(Input) os código do produto
'objProduto RETORNA(Output) os dados do produto

Dim lErro As Long

On Error GoTo Erro_Produto_Obtem_Dados

    'Se o produto foi informado
    If Len(Trim(objProduto.sCodigo)) > 0 Then
    
        'Se o nome reduzido do cliente selecionado não se alterou
        If objProduto.sCodigo = sUltProduto Then
        
            'Significa que o cliente não foi alterado e o código é o mesmo calculado anteriormente
            Set objProduto = objUltProduto
        
        'Senão, significa que o produto foi alterado e é necessário ler novamente
        Else
        
            'Lê os dados do produto a partir do código guardado no obj
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 102231
            
            'Se não encontrou o produto => erro
            If lErro = 28030 Then gError 102232

            'Guarda na variável o código do último produto que teve seus dados
            'lidos no BD
            sUltProduto = objProduto.sCodigo
            
            'Instancia objUltProduto com os dados do último produto lido
            Set objUltProduto = objProduto
        
        End If
    
    'Senão, ou seja, se o produto não foi informado
    Else
    
        'Limpa os dados armazenados do último produto lido
        Call Limpa_Dados_UltProduto
        
    End If

    Produto_Obtem_Dados = SUCESSO
    
    Exit Function

Erro_Produto_Obtem_Dados:

    Produto_Obtem_Dados = gErr
    
    Select Case gErr
    
        Case 102231
        
        Case 102232
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152212)
    
    End Select
    
    Exit Function
    
End Function

'Incluida por Luiz em 30/05/03
Private Sub Limpa_Dados_UltProduto()

    'Limpa a variável que guarda o nome reduzido do último produto que teve
    'seus dados lidos no BD
    sUltProduto = STRING_VAZIO
    
    'Limpa o obj que guarda os dados do último produto lido
    Set objUltProduto = Nothing

End Sub

Public Function Calcula_MnemonicoComissoes_PrecoTabela(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o preço de tabela para cada produto no grid de itens. A Tabela é passada como parametro

Dim iLinha As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUM As String
Dim dFator As Double
Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_Calcula_MnemonicoComissoes_PrecoTabela
    
    objTabelaPrecoItem.iCodTabela = StrParaInt(objMnemonicoValor.vParam(1))
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        'Guarda o produto da linha atual
        sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'formata o produto para que o mesmo possa ser usado como filtro
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 126140
        
        'Se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            'Guarda o código do produto no obj
            objProduto.sCodigo = sProdutoFormatado
            
            'Obtém os dados do produto
            lErro = Produto_Obtem_Dados(objProduto)
            If lErro <> SUCESSO Then gError 126141
            
            'Guarda a U.M. do produto da linha atual
            sUM = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_UnidadeMed_Col)
            
            'Encontra o fator de conversão da UM que está no grid para a UM base de venda
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 126142

            objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

            'pesquisa o preco atual do produto na tabela
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 126143

            'se nao encontrou o preco ==> preco = 0
            If lErro = 28014 Then objTabelaPrecoItem.dPreco = 0

            'Adiciona o preço unitário do produto à coleção
            objMnemonicoValor.colValor.Add objTabelaPrecoItem.dPreco
        
        'Senão
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
        End If
        
    Next

    Calcula_MnemonicoComissoes_PrecoTabela = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_PrecoTabela:

    Calcula_MnemonicoComissoes_PrecoTabela = gErr
    
    Select Case gErr
    
        Case 126140 To 126143
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152213)
    
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_MnemonicoComissoes_Produto_PercDesc(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o % de desconto para cada produto no grid de itens
'objMnemonicoValor RETORNA(Output) o % de desconto encontrado

Dim iLinha As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dPercDesc As Double
Dim lErro As Long, lTamanho As Long

On Error GoTo Erro_Calcula_MnemonicoComissoes_Produto_PercDesc
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        'Guarda o produto da linha atual
        sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'formata o produto para que o mesmo possa ser usado como filtro
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 102240
        
        'Se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            lTamanho = Len(Trim(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PercDesc_Col)))
            If lTamanho > 0 Then
                dPercDesc = CDbl(Format(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PercDesc_Col), "General Number"))
            Else
                dPercDesc = 0
            End If
            
            'Adiciona o preço unitário do produto à coleção
            objMnemonicoValor.colValor.Add dPercDesc
        
        'Senão
        Else
            
            'Guarda na coleção o valor zero que indica que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
        End If
        
    Next

    Calcula_MnemonicoComissoes_Produto_PercDesc = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_Produto_PercDesc:

    Calcula_MnemonicoComissoes_Produto_PercDesc = gErr
    
    Select Case gErr
    
        Case 102240 To 102242
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152202)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_Vend_Vinculo_Cod(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim iVendedor As Integer
Dim iCodVinculo As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Calcula_MnemonicoComissoes_Vend_Vinculo_Cod

    'Alterada por Luiz Nogueira em 22/04/04 para atender às comissões do loja
    'Se estiver calculando comissões do loja, o código do vendedor está na própria tela
    If objTela.Caption = NOME_TELA_COMISSOESCALCULALOJA Then
    
        iVendedor = StrParaInt(objTela.Controls("VendedorCupom").Text)
    
    
    Else
        
        'Obtém os dados da filial do cliente selecionada
        lErro = Obtem_FilialCliente_Dados(objFilialCliente)
        If lErro <> SUCESSO Then gError 94985
    
        'Guarda na coleção o código do vendedor
        iVendedor = objFilialCliente.iVendedor
    
    End If
    
    If iVendedor <> 0 Then
    
        objVendedor.iCodigo = iVendedor
        
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO Then gError 94985
        
        iCodVinculo = objVendedor.iVinculo
    
    End If
    
    objMnemonicoValor.colValor.Add iCodVinculo
    
    Calcula_MnemonicoComissoes_Vend_Vinculo_Cod = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_Vend_Vinculo_Cod:

    Calcula_MnemonicoComissoes_Vend_Vinculo_Cod = gErr
    
    Select Case gErr
    
        Case 94985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_Vend_Vinculo_Desc(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim iVendedor As Integer
Dim iCodVinculo As Integer
Dim sVinculo As String
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Calcula_MnemonicoComissoes_Vend_Vinculo_Desc

    'Alterada por Luiz Nogueira em 22/04/04 para atender às comissões do loja
    'Se estiver calculando comissões do loja, o código do vendedor está na própria tela
    If objTela.Caption = NOME_TELA_COMISSOESCALCULALOJA Then
    
        iVendedor = StrParaInt(objTela.Controls("VendedorCupom").Text)
    
    
    Else
        
        'Obtém os dados da filial do cliente selecionada
        lErro = Obtem_FilialCliente_Dados(objFilialCliente)
        If lErro <> SUCESSO Then gError 94985
    
        'Guarda na coleção o código do vendedor
        iVendedor = objFilialCliente.iVendedor
    
    End If
    
    If iVendedor <> 0 Then
    
        objVendedor.iCodigo = iVendedor
        
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO Then gError 94985
        
        iCodVinculo = objVendedor.iVinculo
        
        '1-Autônomo
        '2-Empregado
        '3-Empresa
        Select Case iCodVinculo
        
            Case 1
                sVinculo = "Autônomo"
            Case 2
                sVinculo = "Empregado"
            Case 3
                sVinculo = "Empresa"
        
        End Select
    
    End If
    
    objMnemonicoValor.colValor.Add sVinculo
    
    Calcula_MnemonicoComissoes_Vend_Vinculo_Desc = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_Vend_Vinculo_Desc:

    Calcula_MnemonicoComissoes_Vend_Vinculo_Desc = gErr
    
    Select Case gErr
    
        Case 94985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TabPreco_Desc(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim iTabela As Integer
Dim sTabela As String
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Calcula_MnemonicoComissoes_TabPreco_Desc

    iTabela = Codigo_Extrai(objTela.Controls("TabelaPreco").Text)

    If iTabela <> 0 Then
    
        objTabelaPreco.iCodigo = iTabela
        
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO Then gError 94985
        
        sTabela = objTabelaPreco.sDescricao
    
    End If
    
    objMnemonicoValor.colValor.Add sTabela
    
    Calcula_MnemonicoComissoes_TabPreco_Desc = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TabPreco_Desc:

    Calcula_MnemonicoComissoes_TabPreco_Desc = gErr
    
    Select Case gErr
    
        Case 94985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TabPreco_Cod(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim iTabela As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_TabPreco_Cod

    iTabela = Codigo_Extrai(objTela.Controls("TabelaPreco").Text)
   
    objMnemonicoValor.colValor.Add iTabela
    
    Calcula_MnemonicoComissoes_TabPreco_Cod = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TabPreco_Cod:

    Calcula_MnemonicoComissoes_TabPreco_Cod = gErr
    
    Select Case gErr
    
        Case 94985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Private Function Existe_objTribTab_Tela() As Boolean

On Error GoTo Erro_Existe_objTribTab_Tela

    Existe_objTribTab_Tela = False
    If Not (objTela.gobjTribTab.gobjDocTribTributacao Is Nothing) Then
        Existe_objTribTab_Tela = True
    End If

    Exit Function

Erro_Existe_objTribTab_Tela:

    Existe_objTribTab_Tela = False

    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_CanalVenda(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iCanalVenda As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_CanalVenda

    iCanalVenda = Codigo_Extrai(objTela.Controls("CanalVenda").Text)
   
    objMnemonicoValor.colValor.Add iCanalVenda
    
    Calcula_MnemonicoComissoes_CanalVenda = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_CanalVenda:

    Calcula_MnemonicoComissoes_CanalVenda = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152199)
    
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_MnemonicoComissoes_ComissaoItem(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o percentual de comissao informado para cada produto no grid de itens (-1 = nao preenchido)

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Calcula_MnemonicoComissoes_ComissaoItem
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        If objTela.iGrid_ComissaoItem_Col <> 0 Then
        
            If Len(Trim(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_ComissaoItem_Col))) = 0 Then
                objMnemonicoValor.colValor.Add -1
            Else
                objMnemonicoValor.colValor.Add PercentParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_ComissaoItem_Col))
            End If
            
        'Senão
        Else
            
            'Guarda na coleção o valor -1 que indica que o campo nao esta preenchido
            objMnemonicoValor.colValor.Add -1
            
        End If
        
    Next

    Calcula_MnemonicoComissoes_ComissaoItem = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_ComissaoItem:

    Calcula_MnemonicoComissoes_ComissaoItem = gErr
    
    Select Case gErr
    
        Case 102240 To 102242
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152202)
    
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_MnemonicoComissoes_TabPrecoItem(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o codigo da tabela de preco informado para cada produto no grid de itens (zero = nao preenchido)

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Calcula_MnemonicoComissoes_TabPrecoItem
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        If objTela.iGrid_TabPrecoItem_Col <> 0 Then
        
            objMnemonicoValor.colValor.Add Codigo_Extrai(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_TabPrecoItem_Col))
            
        'Senão
        Else
            
            'Guarda na coleção o valor zero que indica que o campo nao esta preenchido
            objMnemonicoValor.colValor.Add 0
            
        End If
        
    Next

    Calcula_MnemonicoComissoes_TabPrecoItem = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_TabPrecoItem:

    Calcula_MnemonicoComissoes_TabPrecoItem = gErr
    
    Select Case gErr
    
        Case 102240 To 102242
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152202)
    
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_MnemonicoComissoes_ComisTabPrecoItem(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve o % de comissao associado à tabela de preco dos itens

Dim iLinha As Integer
Dim lErro As Long, iTabPreco As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dComissao As Double
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Calcula_MnemonicoComissoes_ComisTabPrecoItem
    
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
        
        iTabPreco = 0
        dComissao = -1
        
        If objTela.iGrid_TabPrecoItem_Col <> 0 Then
        
             iTabPreco = Codigo_Extrai(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_TabPrecoItem_Col))
            
        End If
        
        If iTabPreco = 0 Then iTabPreco = Codigo_Extrai(objTela.Controls("TabelaPreco").Text)
        
        If iTabPreco <> 0 Then
        
            'Guarda o produto da linha atual
            sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
            
            'formata o produto para que o mesmo possa ser usado como filtro
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 126140
            
            'Se o produto está preenchido
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
                objTabelaPrecoItem.sCodProduto = sProdutoFormatado
                objTabelaPrecoItem.iCodTabela = iTabPreco
                
                'pesquisa o preco atual do produto na tabela
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM

                If lErro = SUCESSO Then dComissao = objTabelaPrecoItem.dComissao

            End If
            
            If dComissao = -1 Then
            
                objTabelaPreco.iCodigo = iTabPreco
                
                lErro = CF("TabelaPreco_Le", objTabelaPreco)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                dComissao = objTabelaPreco.dComissao
            
            End If
            
            objMnemonicoValor.colValor.Add dComissao
            
        Else
        
            objMnemonicoValor.colValor.Add -1
            
        End If
        
    Next

    Calcula_MnemonicoComissoes_ComisTabPrecoItem = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_ComisTabPrecoItem:

    Calcula_MnemonicoComissoes_ComisTabPrecoItem = gErr
    
    Select Case gErr
    
        Case 102240 To 102242, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152202)
    
    End Select
    
    Exit Function
    
End Function


