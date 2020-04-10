VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function FilFornFilEmp_Le_Lock(lComando As Long, objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Lê e lock a tabela FilFornFilEmp com as chaves iFilialEmpresa, lCodFornecedor, iCodFilial vinda do objFilFornFilEmp

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilFornFilEmp_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_FILFORNFILEMP_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 54671
    
    
    FilFornFilEmp_Le_Lock = X.FilFornFilEmp_Le_Lock(lComando, objFilFornFilEmp)

    Exit Function
    
Erro_FilFornFilEmp_Le_Lock:
    
    FilFornFilEmp_Le_Lock = Err
    
    Select Case Err

        Case 54671
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145275)
        
    End Select
    
    Exit Function

End Function

Function TipoDocumento_Le(objTipoDocumento As ClassTipoDocumento) As Long
'Lê o tipo de documento a partir da chave fornecida em objTipoDocumento. Retorna os dados em objTipoDocumento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocumento_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPODOCUMENTO_LE", X)
    If lErro <> SUCESSO Then Error 19886
    
    
    TipoDocumento_Le = X.TipoDocumento_Le(objTipoDocumento)
    
    Exit Function
    
Erro_TipoDocumento_Le:

    TipoDocumento_Le = Err
    
    Select Case Err

        Case 19886
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145276)
        
    End Select
    
    Exit Function
    
End Function

Function MvPerCli_Exclui_Cliente(lCodCliente As Long) As Long
'Exclui os registros da tabela MvPerCli associados ao Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MvPerCli_Exclui_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_MVPERCLI_EXCLUI_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14958
    
    
    MvPerCli_Exclui_Cliente = X.MvPerCli_Exclui_Cliente(lCodCliente)
    
    Exit Function
    
Erro_MvPerCli_Exclui_Cliente:

    MvPerCli_Exclui_Cliente = Err
    
    Select Case Err

        Case 14958
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145277)
        
    End Select
    
    Exit Function
    
End Function

Function MvDiaCli_Exclui_Cliente(lCodCliente As Long) As Long
'Exclui os registros da tabela MvDiaCli associados ao Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MvDiaCli_Exclui_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_MVDIACLI_EXCLUI_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14959
    
    
    MvDiaCli_Exclui_Cliente = X.MvDiaCli_Exclui_Cliente(lCodCliente)
    
    Exit Function
    
Erro_MvDiaCli_Exclui_Cliente:

    MvDiaCli_Exclui_Cliente = Err
    
    Select Case Err

        Case 14959
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145278)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_Comissoes(objNFiscal As ClassNFiscal) As Long
'preenche a colecao de comissoes associadas a uma NF.

Dim X As Object, lErro As Long

On Error GoTo Erro_NFiscal_Le_Comissoes

    
    

    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFSREC_LE_COMISSOES", X)
    If lErro <> SUCESSO Then Error 24566

    
    NFiscal_Le_Comissoes = X.NFiscal_Le_Comissoes(objNFiscal)

    Exit Function

Erro_NFiscal_Le_Comissoes:

    NFiscal_Le_Comissoes = Err

    Select Case Err

        Case 24566
            

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145279)

    End Select

    Exit Function

End Function

Function Comissoes_Grava(colComissao As colComissao) As Long
'Grava as Comissões passadas em colComissao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_GRAVA", X)
    If lErro <> SUCESSO Then Error 24567
    
    
    Comissoes_Grava = X.Comissoes_Grava(colComissao)

    Exit Function
    
Erro_Comissoes_Grava:
    
    Comissoes_Grava = Err
    
    Select Case Err

        Case 24567
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145280)
        
    End Select
    
    Exit Function

End Function

Function Comissao_Automatico(lNumIntDoc As Long) As Long
'Gera o próximo Número Interno de uma Comissao
'AVISO: ESSA ROTINA DEVE SER CHAMADA DE DENTRO DE UMA TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissao_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSAO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 24568
    
    
    Comissao_Automatico = X.Comissao_Automatico(lNumIntDoc)

    Exit Function
    
Erro_Comissao_Automatico:
    
    Comissao_Automatico = Err
    
    Select Case Err

        Case 24568
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145281)
        
    End Select
    
    Exit Function

End Function
Function PorcentagemJuros_Critica(sString As String) As Long
'Critica o valor passado em sString , se nao for entre 0 e PORCENTAGEM_JUROS_MAXIMA retorna erro.
Dim X As Object, lErro As Long
    
On Error GoTo Erro_PorcentagemJuros_Critica
    

    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PORCENTAGEM_JUROS_CRITICA", X)
    If lErro <> SUCESSO Then Error 25090
    
    
    PorcentagemJuros_Critica = X.PorcentagemJuros_Critica(sString)

    Exit Function
    
Erro_PorcentagemJuros_Critica:
    
    PorcentagemJuros_Critica = Err
    
    Select Case Err

        Case 25090
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145282)
        
    End Select

End Function

Function Estados_Le_Todos(colEstados As Collection) As Long
'carrega a tabela de estados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estados_Le_Todos
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_ESTADOS_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 24596
    
    
    Estados_Le_Todos = X.Estados_Le_Todos(colEstados)

    Exit Function
    
Erro_Estados_Le_Todos:
    
    Estados_Le_Todos = Err
    
    Select Case Err

        Case 24596
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145283)
        
    End Select

End Function

Function TipoDocInfo_Le(objTipoDocInfo As ClassTipoDocInfo) As Long
'objTipoDocInfo.sSigla tem que estar preenchido
'Preenche objTipoDocInfo a partir do bd p/registro identificado por objTipoDocInfo.sSigla
'Retorna 27263 se nao encontrar

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPODOCINFO_LE", X)
    If lErro <> SUCESSO Then Error 24597
    
    
    TipoDocInfo_Le = X.TipoDocInfo_Le(objTipoDocInfo)

    Exit Function
    
Erro_TipoDocInfo_Le:
    
    TipoDocInfo_Le = Err
    
    Select Case Err

        Case 24597
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145284)
        
    End Select

End Function

Function Vendedor_Le(objVendedor As ClassVendedor) As Long
'Le todos os dados da tabela Vendedores, correspondente ao codigo do vendedor passado por parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_VENDEDOR_LE", X)
    If lErro <> SUCESSO Then Error 14332
    
    
    Vendedor_Le = X.Vendedor_Le(objVendedor)
    
    Exit Function
    
Erro_Vendedor_Le:

    Vendedor_Le = Err
    
    Select Case Err

        Case 14332
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145285)
        
    End Select
    
    Exit Function
    
End Function

Function Vendedor_Le_NomeReduzido(objVendedor As ClassVendedor) As Long
'Lê os dados do Vendedor, correspondente ao NomeReduzido passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Le_NomeReduzido
    


    
    lErro = Rotina_ObterObjetoFuncao("CPR_VENDEDOR_LE_NOME_REDUZIDO", X)
    If lErro <> SUCESSO Then Error 25009
    
    
    Vendedor_Le_NomeReduzido = X.Vendedor_Le_NomeReduzido(objVendedor)
    
    Exit Function
    
Erro_Vendedor_Le_NomeReduzido:

    Vendedor_Le_NomeReduzido = Err
    
    Select Case Err

        Case 25009
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145286)
        
    End Select

End Function

Function Cliente_Le_NomeReduzido(objCliente As ClassCliente) As Long
'le o cliente a partir da chave(sNomeReduzido) fornecida em objCliente. Devolve os dados em objCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Le_NomeReduzido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_LE_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then Error 14323
    
    
    Cliente_Le_NomeReduzido = X.Cliente_Le_NomeReduzido(objCliente)
    
    Exit Function
    
Erro_Cliente_Le_NomeReduzido:

    Cliente_Le_NomeReduzido = Err
    
    Select Case Err

        Case 14323
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145287)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_Le(objCliente As ClassCliente) As Long
'le o cliente a partir da chave(lCodigo) fornecida em objCliente. Devolve os dados em objCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_LE", X)
    If lErro <> SUCESSO Then Error 14322
    
    
    Cliente_Le = X.Cliente_Le(objCliente)
    
    Exit Function
    
Erro_Cliente_Le:

    Cliente_Le = Err
    
    Select Case Err

        Case 14322
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145288)
        
    End Select
    
    Exit Function
    
End Function
Function Cliente_Le_Cgc(objCliente As ClassCliente, iCodFilial As Integer) As Long
'le o Cliente a partir do Cgc fornecido em objCliente. Devolve os dados em objCliente.
'Devolve o código da filial em iCodFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Le_Cgc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_LE_CGC", X)
    If lErro <> SUCESSO Then Error 6724
    
    
    Cliente_Le_Cgc = X.Cliente_Le_Cgc(objCliente, iCodFilial)
    
    Exit Function
    
Erro_Cliente_Le_Cgc:

    Cliente_Le_Cgc = Err
    
    Select Case Err

        Case 6724
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145289)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_Le_NomeRed(objCodigoNome As AdmlCodigoNome) As Long
'Le o NomeReduzido do cliente a partir da chave(lCodigo) fornecida em objCodigoNome.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Le_NomeRed
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_LE_NOMERED", X)
    If lErro <> SUCESSO Then Error 14341
    
    
    Cliente_Le_NomeRed = X.Cliente_Le_NomeRed(objCodigoNome)
    
    Exit Function
    
Erro_Cliente_Le_NomeRed:

    Cliente_Le_NomeRed = Err
    
    Select Case Err

        Case 14341
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145290)
        
    End Select
    
    Exit Function
    
End Function

''Function Endereco_Le(objEndereco As ClassEndereco) As Long
'''le o endereco a partir da chave fornecida em objEndereco. Devolve os dados em objEndereco
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Endereco_Le
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_LE",X)
''    If lErro <> SUCESSO Then Error 14325
''
''
''    Endereco_Le = X.Endereco_Le(objEndereco)
''
''    Exit Function
''
''Erro_Endereco_Le:
''
''    Endereco_Le = Err
''
''    Select Case Err
''
''        Case 14325
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145291)
''
''    End Select
''
''    Exit Function
''
''End Function
''
Function RegiaoVenda_Exclui(objRegiaoVenda As ClassRegiaoVenda) As Long
'exclui regiao de venda do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegiaoVenda_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_REGIAOVENDA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14293
    
    
    RegiaoVenda_Exclui = X.RegiaoVenda_Exclui(objRegiaoVenda)
    
    Exit Function
    
Erro_RegiaoVenda_Exclui:

    RegiaoVenda_Exclui = Err
    
    Select Case Err

        Case 14293
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145292)
        
    End Select
    
    Exit Function
    
End Function

Function RegiaoVenda_Grava(objRegiaoVenda As ClassRegiaoVenda) As Long
'Atualiza ou insere um novo registro na tabela RegioesVendas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegiaoVenda_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_REGIAOVENDA_GRAVA", X)
    If lErro <> SUCESSO Then Error 14294
    
    
    RegiaoVenda_Grava = X.RegiaoVenda_Grava(objRegiaoVenda)
    
    Exit Function
    
Erro_RegiaoVenda_Grava:

    RegiaoVenda_Grava = Err
    
    Select Case Err

        Case 14294
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145293)
        
    End Select
    
    Exit Function
    
End Function

Function RegiaoVenda_Automatico(iCodigo As Integer) As Long
'retorna o código da próxima Regiao de Venda disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegiaoVenda_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_REGIAOVENDA_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 14295
    
    
    RegiaoVenda_Automatico = X.RegiaoVenda_Automatico(iCodigo)
    
    Exit Function
    
Erro_RegiaoVenda_Automatico:

    RegiaoVenda_Automatico = Err
    
    Select Case Err

        Case 14295
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145294)
        
    End Select
    
    Exit Function
    
End Function

Function RegiaoVenda_Le(objRegiaoVenda As ClassRegiaoVenda) As Long
'le os dados da Regiao de Venda, cujo codigo foi passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegiaoVenda_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_REGIAOVENDA_LE", X)
    If lErro <> SUCESSO Then Error 14296
    
    
    RegiaoVenda_Le = X.RegiaoVenda_Le(objRegiaoVenda)
    
    Exit Function
    
Erro_RegiaoVenda_Le:

    RegiaoVenda_Le = Err
    
    Select Case Err

        Case 14296
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145295)
        
    End Select
    
    Exit Function
    
End Function

Function Vendedores_Le_Regiao(iCodigo As Integer) As Long
'le os vendedores relacionados ao codigo da regiao de venda passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedores_Le_Regiao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_VENDEDORES_LE_REGIAO", X)
    If lErro <> SUCESSO Then Error 14297
    
    
    Vendedores_Le_Regiao = X.Vendedores_Le_Regiao(iCodigo)
    
    Exit Function
    
Erro_Vendedores_Le_Regiao:

    Vendedores_Le_Regiao = Err
    
    Select Case Err

        Case 14297
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145296)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDeCliente_Le_Regiao(iCodigo As Integer) As Long
'le os tipos de clientes relacionados ao codigo da regiao de venda passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDeCliente_Le_Regiao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOSDECLIENTE_LE_REGIAO", X)
    If lErro <> SUCESSO Then Error 14298
    
    
    TiposDeCliente_Le_Regiao = X.TiposDeCliente_Le_Regiao(iCodigo)
    
    Exit Function
    
Erro_TiposDeCliente_Le_Regiao:

    TiposDeCliente_Le_Regiao = Err
    
    Select Case Err

        Case 14298
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145297)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Le_Regiao(iCodigo As Integer) As Long
'le as filiais clientes relacionadas ao codigo da regiao de venda passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Le_Regiao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_LE_REGIAO", X)
    If lErro <> SUCESSO Then Error 14299
    
    
    FiliaisClientes_Le_Regiao = X.FiliaisClientes_Le_Regiao(iCodigo)
    
    Exit Function
    
Erro_FiliaisClientes_Le_Regiao:

    FiliaisClientes_Le_Regiao = Err
    
    Select Case Err

        Case 14299
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145298)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Le_Tabelas_Lock(objFilialCliente As ClassFilialCliente) As Long
'Faz o lock das tabelas relacionadas a tabela FiliaisClientes para a gravação de novo Cliente
'CHAMADA EM TRANSAÇÃO por FiliaisClientes_Grava

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Le_Tabelas_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_LE_TABELAS_LOCK", X)
    If lErro <> SUCESSO Then Error 14338
    
    
    FiliaisClientes_Le_Tabelas_Lock = X.FiliaisClientes_Le_Tabelas_Lock(objFilialCliente)
    
    Exit Function
    
Erro_FiliaisClientes_Le_Tabelas_Lock:

    FiliaisClientes_Le_Tabelas_Lock = Err
    
    Select Case Err

        Case 14338
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145299)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Le_Chave(colFiliais As ColChaveFilial) As Long
'Le todos os Codigos de Cliente , Codigos de Filiais e Nomes da tabela Filiais de Clientes e coloca na colecao ColFiliais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Le_Chave
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_LE_CHAVE", X)
    If lErro <> SUCESSO Then Error 14339
    
    
    FiliaisClientes_Le_Chave = X.FiliaisClientes_Le_Chave(colFiliais)
    
    Exit Function
    
Erro_FiliaisClientes_Le_Chave:

    FiliaisClientes_Le_Chave = Err
    
    Select Case Err

        Case 14339
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145300)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Le(objFilialCliente As ClassFilialCliente) As Long
'Le todos os dados da tabela FiliaisClientes, correspondente a chave passada por parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIALCLIENTE_LE", X)
    If lErro <> SUCESSO Then Error 14340
    
    
    FilialCliente_Le = X.FilialCliente_Le(objFilialCliente)
    
    Exit Function
    
Erro_FilialCliente_Le:

    FilialCliente_Le = Err
    
    Select Case Err

        Case 14340
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145301)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_Automatico(lCodigo As Long) As Long
'retorna o número do proximo Cliente disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 14316
    
    
    Cliente_Automatico = X.Cliente_Automatico(lCodigo)
    
    Exit Function
    
Erro_Cliente_Automatico:

    Cliente_Automatico = Err
    
    Select Case Err

        Case 14316
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145302)
        
    End Select
    
    Exit Function
    
End Function

''Function Endereco_Automatico(lCodigo As Long) As Long
'''retorna o número do proximo Endereco disponivel
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Endereco_Automatico
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_AUTOMATICO",X)
''    If lErro <> SUCESSO Then Error 14317
''
''
''    Endereco_Automatico = X.Endereco_Automatico(lCodigo)
''
''    Exit Function
''
''Erro_Endereco_Automatico:
''
''    Endereco_Automatico = Err
''
''    Select Case Err
''
''        Case 14317
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145303)
''
''    End Select
''
''    Exit Function
''
''End Function
''
Function Cliente_Grava(objCliente As ClassCliente, colEndereco As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_GRAVA", X)
    If lErro <> SUCESSO Then Error 14318
    
    
    Cliente_Grava = X.Cliente_Grava(objCliente, colEndereco)
    
    Exit Function
    
Erro_Cliente_Grava:

    Cliente_Grava = Err
    
    Select Case Err

        Case 14318
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145304)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_Exclui(objCliente As ClassCliente) As Long
'Exclui o cliente , suas filiais e seus Enderecos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14319
    
    
    Cliente_Exclui = X.Cliente_Exclui(objCliente)
    
    Exit Function
    
Erro_Cliente_Exclui:

    Cliente_Exclui = Err
    
    Select Case Err

        Case 14319
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145305)
        
    End Select
    
    Exit Function
    
End Function

''Function Endereco_Grava(objEndereco As ClassEndereco) As Long
'''grava o endereco no banco de dados
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Endereco_Grava
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_GRAVA",X)
''    If lErro <> SUCESSO Then Error 14320
''
''
''    Endereco_Grava = X.Endereco_Grava(objEndereco)
''
''    Exit Function
''
''Erro_Endereco_Grava:
''
''    Endereco_Grava = Err
''
''    Select Case Err
''
''        Case 14320
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145306)
''
''    End Select
''
''    Exit Function
''
''End Function
''
Function Enderecos_Le_Cliente(colEnderecos As colEndereco, objCliente As ClassCliente) As Long
'Le cada um dos 3 tipos de enderecos(principal,entrega e cobranca) na tabela Enderecos e coloca na colecao

Dim X As Object, lErro As Long

On Error GoTo Erro_Enderecos_Le_Cliente

    
    

    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECOS_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14324

    
    Enderecos_Le_Cliente = X.Enderecos_Le_Cliente(colEnderecos, objCliente)

    Exit Function

Erro_Enderecos_Le_Cliente:

    Enderecos_Le_Cliente = Err

    Select Case Err

        Case 14324
            

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145307)

    End Select

    Exit Function

End Function
''
''Function Endereco_Exclui(lCodEndereco As Long) As Long
'''Exclui o Endereco cujo código foi passado
'''Chamada DENTRO DE TRANSAÇÃO
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Endereco_Exclui
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_EXCLUI",X)
''    If lErro <> SUCESSO Then Error 14348
''
''
''    Endereco_Exclui = X.Endereco_Exclui(lCodEndereco)
''
''    Exit Function
''
''Erro_Endereco_Exclui:
''
''    Endereco_Exclui = Err
''
''    Select Case Err
''
''        Case 14348
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145308)
''
''    End Select
''
''    Exit Function
''
''End Function
''
Function TipoCliente_Le(objTipoCliente As ClassTipoCliente) As Long
'le os dados do Tipo de Cliente a partir da chave . Devolve os dados em objTipoCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOCLIENTE_LE", X)
    If lErro <> SUCESSO Then Error 14326
    
    
    TipoCliente_Le = X.TipoCliente_Le(objTipoCliente)
    
    Exit Function
    
Erro_TipoCliente_Le:

    TipoCliente_Le = Err
    
    Select Case Err

        Case 14326
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145309)
        
    End Select
    
    Exit Function
    
End Function

Function Clientes_Le_Tabelas_Lock(objCliente As ClassCliente) As Long
'Faz o lock das tabelas relacionadas a tabela Cliente para assegurar a gravação de um novo Cliente
'Verifica se Condição de Pagto pode ser usada em Contas Receber
'Chamada DENTRO DE TRANSAÇÃO por Cliente_Grava

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Clientes_Le_Tabelas_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTES_LE_TABELAS_LOCK", X)
    If lErro <> SUCESSO Then Error 14333
    
    
    Clientes_Le_Tabelas_Lock = X.Clientes_Le_Tabelas_Lock(objCliente)
    
    Exit Function
    
Erro_Clientes_Le_Tabelas_Lock:

    Clientes_Le_Tabelas_Lock = Err
    
    Select Case Err

        Case 14333
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145310)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Lock(lCodigo As Long) As Long
'faz o lock das Filiais Clientes relacionadas ao cliente passado por parametro para garantir a exclusao de um Cliente
'funcao chamada por Cliente_Exclui dentro de uma transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_LOCK", X)
    If lErro <> SUCESSO Then Error 14334
    
    
    FiliaisClientes_Lock = X.FiliaisClientes_Lock(lCodigo)
    
    Exit Function
    
Erro_FiliaisClientes_Lock:

    FiliaisClientes_Lock = Err
    
    Select Case Err

        Case 14334
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145311)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Exclui(objFilialCliente As ClassFilialCliente) As Long
'Exclui a Filial Cliente e seus Enderecos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALCLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14346
    
    
    FilialCliente_Exclui = X.FilialCliente_Exclui(objFilialCliente)
    
    Exit Function
    
Erro_FilialCliente_Exclui:

    FilialCliente_Exclui = Err
    
    Select Case Err

        Case 14346
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145312)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Automatico(ByVal lCodCliente As Long, iCodigo As Integer) As Long
'retorna o número da proxima Filial Cliente disponivel em iCodigo em relacao ao numero de cliente passado em lCodigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 14347
    
    
    FiliaisClientes_Automatico = X.FiliaisClientes_Automatico(lCodCliente, iCodigo)
    
    Exit Function
    
Erro_FiliaisClientes_Automatico:

    FiliaisClientes_Automatico = Err
    
    Select Case Err

        Case 14347
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145313)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Grava(objFilialCliente As ClassFilialCliente, colEndereco As Collection) As Long
'Grava a Filial Cliente e seus 3 tipos de Enderecos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_GRAVA", X)
    If lErro <> SUCESSO Then Error 14349
    
    
    FiliaisClientes_Grava = X.FiliaisClientes_Grava(objFilialCliente, colEndereco)
    
    Exit Function
    
Erro_FiliaisClientes_Grava:

    FiliaisClientes_Grava = Err
    
    Select Case Err

        Case 14349
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145314)
        
    End Select
    
    Exit Function
    
End Function

Function Enderecos_Exclui_Cliente(objCliente As ClassCliente) As Long
'Exclui os Endereços do Cliente (principal, entrega e cobrança)
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim X As Object, lErro As Long

On Error GoTo Erro_Enderecos_Exclui_Cliente

    
    

    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECOS_EXCLUI_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14960

    
    Enderecos_Exclui_Cliente = X.Enderecos_Exclui_Cliente(objCliente)

    Exit Function

Erro_Enderecos_Exclui_Cliente:

    Enderecos_Exclui_Cliente = Err

    Select Case Err

        Case 14960
            

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145315)

    End Select

    Exit Function

End Function

Function FiliaisClientes_Exclui(lCodCliente As Long) As Long
'Exclui as Filiais do Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14961
    
    
    FiliaisClientes_Exclui = X.FiliaisClientes_Exclui(lCodCliente)
    
    Exit Function
    
Erro_FiliaisClientes_Exclui:

    FiliaisClientes_Exclui = Err
    
    Select Case Err

        Case 14961
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145316)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_ObterDadosBasicos(objCliente As ClassCliente) As Long
'le o cliente a partir da chave(lCodigo) fornecida em objCliente.
'Devolve os dados em objCliente
'nao acessa dados da matriz do cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_ObterDadosBasicos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_OBTERDADOSBASICOS", X)
    If lErro <> SUCESSO Then Error 14969
    
    
    Cliente_ObterDadosBasicos = X.Cliente_ObterDadosBasicos(objCliente)
    
    Exit Function
    
Erro_Cliente_ObterDadosBasicos:

    Cliente_ObterDadosBasicos = Err
    
    Select Case Err

        Case 14969
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145317)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Le_NomeRed_CodFilial(sNomeRed As String, objFilialCliente As ClassFilialCliente) As Long
'Le a FilialClliente a partir do NomeReduzido do Cliente passado em sNomeRed
'Le do código da Filial passado em objFilialCliente. Devolve a Filial em objFilialCiente.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Le_NomeRed_CodFilial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILCLI_LE_NOMERED_CODFILIAL", X)
    If lErro <> SUCESSO Then Error 27009
    
    
    FilialCliente_Le_NomeRed_CodFilial = X.FilialCliente_Le_NomeRed_CodFilial(sNomeRed, objFilialCliente)
    
    Exit Function
    
Erro_FilialCliente_Le_NomeRed_CodFilial:

    FilialCliente_Le_NomeRed_CodFilial = Err
    
    Select Case Err

        Case 27009
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145318)
        
    End Select
    
    Exit Function
    
End Function

Function FiliaisClientes_Le_Cliente(objCliente As ClassCliente, colCodigoNome As AdmColCodigoNome) As Long
'Le na tabela FiliaisClientes todos os Codigos e Nomes de Filiais
'relacionadas ao objCliente. Retorna na colecao colCodigoNome

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisClientes_Le_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISCLIENTES_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 6726
    
    
    FiliaisClientes_Le_Cliente = X.FiliaisClientes_Le_Cliente(objCliente, colCodigoNome)
    
    Exit Function
    
Erro_FiliaisClientes_Le_Cliente:

    FiliaisClientes_Le_Cliente = Err
    
    Select Case Err

        Case 6726
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145319)
        
    End Select
    
    Exit Function
    
End Function

Function Enderecos_Exclui_FilialCliente(objFilialCliente As ClassFilialCliente) As Long
'Exclui os  Enderecos Principal, Entrega e Cobranca da Filial Cliente
'CHAMADA EM TRANSAÇÃO por FilialCliente_Exclui e por FiliaisClientes_Exclui

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Enderecos_Exclui_FilialCliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECOS_EXCLUI_FILIALCLIENTE", X)
    If lErro <> SUCESSO Then Error 24024
    
    
    Enderecos_Exclui_FilialCliente = X.Enderecos_Exclui_FilialCliente(objFilialCliente)
    
    Exit Function
    
Erro_Enderecos_Exclui_FilialCliente:

    Enderecos_Exclui_FilialCliente = Err
    
    Select Case Err

        Case 24024
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145320)
        
    End Select
    
    Exit Function
    
End Function

Function TipoCliente_Automatico(iCodigo As Integer) As Long
'retorna o número do proximo TipoCliente disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOCLIENTE_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 14485
    
    
    TipoCliente_Automatico = X.TipoCliente_Automatico(iCodigo)
    
    Exit Function
    
Erro_TipoCliente_Automatico:

    TipoCliente_Automatico = Err
    
    Select Case Err

        Case 14485
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145321)
        
    End Select
    
    Exit Function
    
End Function

Function TipoCliente_Grava(objTipoCliente As ClassTipoCliente) As Long
'Grava o Tipo Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOCLIENTE_GRAVA", X)
    If lErro <> SUCESSO Then Error 14486
    
    
    TipoCliente_Grava = X.TipoCliente_Grava(objTipoCliente)
    
    Exit Function
    
Erro_TipoCliente_Grava:

    TipoCliente_Grava = Err
    
    Select Case Err

        Case 14486
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145322)
        
    End Select
    
    Exit Function
    
End Function

Function TipoCliente_Exclui(objTipoCliente As ClassTipoCliente) As Long
'Exclui o TipoCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOCLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14487
    
    
    TipoCliente_Exclui = X.TipoCliente_Exclui(objTipoCliente)
    
    Exit Function
    
Erro_TipoCliente_Exclui:

    TipoCliente_Exclui = Err
    
    Select Case Err

        Case 14487
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145323)
        
    End Select
    
    Exit Function
    
End Function

Function TipoCliente_Le_Tabelas_Lock(objTipoCliente As ClassTipoCliente) As Long
'faz o lock das tabelas relacionadas a tabela TiposDeCliente para assegurar a gravacao de um novo Tipo de Cliente
'funcao chamada por TipoCliente_Grava dentro de uma transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Le_Tabelas_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TIPOCLIENTE_LE_TABELAS_LOCK", X)
    If lErro <> SUCESSO Then Error 14484
    
    
    TipoCliente_Le_Tabelas_Lock = X.TipoCliente_Le_Tabelas_Lock(objTipoCliente)
    
    Exit Function
    
Erro_TipoCliente_Le_Tabelas_Lock:

    TipoCliente_Le_Tabelas_Lock = Err
    
    Select Case Err

        Case 14484
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145324)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoCobranca_ObterNaColecao(objPadraoCobranca As ClassPadraoCobranca, colPadroesCobranca As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_ObterNaColecao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOCOBRANCA_OBTERNACOLECAO", X)
    If lErro <> SUCESSO Then Error 24704
    
    
    PadraoCobranca_ObterNaColecao = X.PadraoCobranca_ObterNaColecao(objPadraoCobranca, colPadroesCobranca)
    
    Exit Function
    
Erro_PadraoCobranca_ObterNaColecao:

    PadraoCobranca_ObterNaColecao = Err
    
    Select Case Err

        Case 24704
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145325)
        
    End Select
    
    Exit Function
    
End Function


Function PadraoCobranca_Le(objPadraoCobranca As ClassPadraoCobranca) As Long
'le o PadraoCobranca a partir da chave fornecida em objPadraoCobranca. Devolve os dados em objPadraoCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_PADRAOCOBRANCA_LE", X)
    If lErro <> SUCESSO Then Error 14586
    
    
    PadraoCobranca_Le = X.PadraoCobranca_Le(objPadraoCobranca)
    
    Exit Function
    
Erro_PadraoCobranca_Le:

    PadraoCobranca_Le = Err
    
    Select Case Err

        Case 14586
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145326)
        
    End Select
    
    Exit Function
    
End Function

Function PadroesCobranca_Le_Todos(colPadroesCobranca As Collection)
'carrega uma colecao com elementos objPadraoCobranca correspondentes
'aos registros ATIVOS E INATIVOS da tabela PadroesCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadroesCobranca_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_PADROESCOBRANCA_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 14966
    
    
    PadroesCobranca_Le_Todos = X.PadroesCobranca_Le_Todos(colPadroesCobranca)
    
    Exit Function
    
Erro_PadroesCobranca_Le_Todos:

    PadroesCobranca_Le_Todos = Err
    
    Select Case Err

        Case 14966
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145327)
        
    End Select
    
    Exit Function
    
End Function

Function NFs_Cliente_Nao_Relacionado(lCodCliente As Long) As Long
'Verifica se Cliente não está relacionado com Notas Fiscais Receber
'Retorna SUCESSO se não estiver relacionado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFs_Cliente_Nao_Relacionado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_NFS_CLIENTE_NAO_RELACIONADO", X)
    If lErro <> SUCESSO Then Error 14956
    
    
    NFs_Cliente_Nao_Relacionado = X.NFs_Cliente_Nao_Relacionado(lCodCliente)
    
    Exit Function
    
Erro_NFs_Cliente_Nao_Relacionado:

    NFs_Cliente_Nao_Relacionado = Err
    
    Select Case Err

        Case 14956
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145328)
        
    End Select
    
    Exit Function
    
End Function

Function TitulosReceber_Cliente_Nao_Relacionado(lCodCliente As Long) As Long
'Verifica se Cliente não está relacionado com Títulos a Receber
'Retorna SUCESSO se não estiver relacionado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TitulosReceber_Cliente_Nao_Relacionado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_TITULOSRECEBER_CLIENTE_NAO_RELACIONA", X)
    If lErro <> SUCESSO Then Error 14957
    
    
    TitulosReceber_Cliente_Nao_Relacionado = X.TitulosReceber_Cliente_Nao_Relacionado(lCodCliente)
    
    Exit Function
    
Erro_TitulosReceber_Cliente_Nao_Relacionado:

    TitulosReceber_Cliente_Nao_Relacionado = Err
    
    Select Case Err

        Case 14957
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145329)
        
    End Select
    
    Exit Function
    
End Function

Function DebitosRecCli_Le_Cliente(lCodigo As Long) As Long
'Le todos códigos internos de Debitos Receber relacionadas ao Cliente
'Devolve em colCodDebitos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DebitosRecCli_Le_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_DEBITOSRECCLI_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14329
    
    
    DebitosRecCli_Le_Cliente = X.DebitosRecCli_Le_Cliente(lCodigo)
    
    Exit Function
    
Erro_DebitosRecCli_Le_Cliente:

    DebitosRecCli_Le_Cliente = Err
    
    Select Case Err

        Case 14329
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145330)
        
    End Select
    
    Exit Function
    
End Function

Function RecebAntecipados_Le_Cliente(lCodigo As Long) As Long
'Le todos códigos internos de Recebimentos Antecipados relacionados a Cliente
'Devolve em colCodRecAntec

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebAntecipados_Le_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_RECEBANTECIPADOS_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14330
    
    
    RecebAntecipados_Le_Cliente = X.RecebAntecipados_Le_Cliente(lCodigo)
    
    Exit Function
    
Erro_RecebAntecipados_Le_Cliente:

    RecebAntecipados_Le_Cliente = Err
    
    Select Case Err

        Case 14330
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145331)
        
    End Select
    
    Exit Function
    
End Function

Function TitulosReceber_Le_Cliente(lCodigo As Long, colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber relacionados a chave do Cliente (lcodigo) passada
'Devolve em colCodTitulosReceber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TitulosReceber_Le_Cliente
    


    
    lErro = Rotina_ObterObjetoFuncao("CPR_TITULOSRECEBER_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 14328
    
    
    TitulosReceber_Le_Cliente = X.TitulosReceber_Le_Cliente(lCodigo, colCodTitulosReceber)
    
    Exit Function
    
Erro_TitulosReceber_Le_Cliente:

    TitulosReceber_Le_Cliente = Err
    
    Select Case Err

        Case 14328
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145332)
        
    End Select
    
    Exit Function
    
End Function

Function TitulosReceberBaixados_Le_Cliente(lCodigo As Long, colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber Baixados relacionados a chave do Cliente (lcodigo) passada
'Devolve em colCodTitulosReceber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TitulosReceberBaixados_Le_Cliente
    


    
    lErro = Rotina_ObterObjetoFuncao("CPR_TITULOSRECEBERBAIXADOS_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 6862
    
    
    TitulosReceberBaixados_Le_Cliente = X.TitulosReceberBaixados_Le_Cliente(lCodigo, colCodTitulosReceber)
    
    Exit Function
    
Erro_TitulosReceberBaixados_Le_Cliente:

    TitulosReceberBaixados_Le_Cliente = Err
    
    Select Case Err

        Case 6862
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145333)
        
    End Select
    
    Exit Function
    
End Function

Function Estado_Le(objEstado As ClassEstado) As Long
'Lê o Estado a partir do código(sigla) em objEstado.
'Devolve os dados em objEstado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estado_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_ESTADO_LE", X)
    If lErro <> SUCESSO Then Error 24878
    
    
    Estado_Le = X.Estado_Le(objEstado)
    
    Exit Function
    
Erro_Estado_Le:

    Estado_Le = Err
    
    Select Case Err

        Case 24878
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145334)
        
    End Select
    
    Exit Function
    
End Function

Function Estado_Grava(objEstado As ClassEstado) As Long
'Grava as alíquotas do Estado no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estado_Grava
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_ESTADO_GRAVA", X)
    If lErro <> SUCESSO Then Error 24879
    
    
    Estado_Grava = X.Estado_Grava(objEstado)
    
    Exit Function
    
Erro_Estado_Grava:

    Estado_Grava = Err
    
    Select Case Err

        Case 24879
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145335)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaCliente_Le_Todos(colCategorias As Collection) As Long
'Lê na tabela CategoriaCliente todas as Categorias, passando para a coleção colCategorias

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaCliente_Le_Todos
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 24880
    
    
    CategoriaCliente_Le_Todos = X.CategoriaCliente_Le_Todos(colCategorias)
    
    Exit Function
    
Erro_CategoriaCliente_Le_Todos:

    CategoriaCliente_Le_Todos = Err
    
    Select Case Err

        Case 24880
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145336)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaCliente_Le(objCategoriaCliente As ClassCategoriaCliente) As Long
'Lê na tabela CategoriaCliente a descrição da categoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaCliente_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_LE", X)
    If lErro <> SUCESSO Then Error 24881
    
    
    CategoriaCliente_Le = X.CategoriaCliente_Le(objCategoriaCliente)
    
    Exit Function
    
Erro_CategoriaCliente_Le:

    CategoriaCliente_Le = Err
    
    Select Case Err

        Case 24881
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145337)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaCliente_Le_Itens(objCategoriaCliente As ClassCategoriaCliente, colItensCategoria As Collection) As Long
'Lê na tabela de CategoriaClienteItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaCliente_Le_Itens
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_LE_ITENS", X)
    If lErro <> SUCESSO Then Error 24882
    
    
    CategoriaCliente_Le_Itens = X.CategoriaCliente_Le_Itens(objCategoriaCliente, colItensCategoria)
    
    Exit Function
    
Erro_CategoriaCliente_Le_Itens:

    CategoriaCliente_Le_Itens = Err
    
    Select Case Err

        Case 24882
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145338)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaCliente_Grava(objCategoriaCliente As ClassCategoriaCliente, colItensCategoria As Collection) As Long
'Inclui ou altera uma categoria de clientes e seus valores
'tabelas: CategoriaCliente e CategoriaClienteItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaCliente_Grava
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_GRAVA", X)
    If lErro <> SUCESSO Then Error 24883
    
    
    CategoriaCliente_Grava = X.CategoriaCliente_Grava(objCategoriaCliente, colItensCategoria)
    
    Exit Function
    
Erro_CategoriaCliente_Grava:

    CategoriaCliente_Grava = Err
    
    Select Case Err

        Case 24883
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145339)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaCliente_Exclui(objCategoriaCliente As ClassCategoriaCliente) As Long
'Excluir a tabela com todos os seus itens se não for usada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaCliente_Exclui
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24884
    
    
    CategoriaCliente_Exclui = X.CategoriaCliente_Exclui(objCategoriaCliente)
    
    Exit Function
    
Erro_CategoriaCliente_Exclui:

    CategoriaCliente_Exclui = Err
    
    Select Case Err

        Case 24884
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145340)
        
    End Select
    
    Exit Function
    
End Function

Function TipoDeCliente_Le(objTipoCliente As ClassTipoCliente) As Long
'Lê a tabela TiposDeCliente à partir do Tipo De Cliente passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeCliente_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSDECLIENTE_LE", X)
    If lErro <> SUCESSO Then Error 24887
    
    
    TipoDeCliente_Le = X.TipoDeCliente_Le(objTipoCliente)
    
    Exit Function
    
Erro_TipoDeCliente_Le:

    TipoDeCliente_Le = Err
    
    Select Case Err

        Case 24887
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145341)
        
    End Select
    
    Exit Function
    
End Function

Function TipoDeClienteCategorias_Le(objTipoCliente As ClassTipoCliente, colCategoriaItem As Collection) As Long
'Lê na tabela de TiposDeClienteCategoria, todas as categorias de um Tipo de Cliente, guardando na coleção colCategoriaItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeClienteCategorias_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPODECLIENTECATEGORIAS_LE", X)
    If lErro <> SUCESSO Then Error 24888
    
    
    TipoDeClienteCategorias_Le = X.TipoDeClienteCategorias_Le(objTipoCliente, colCategoriaItem)
    
    Exit Function
    
Erro_TipoDeClienteCategorias_Le:

    TipoDeClienteCategorias_Le = Err
    
    Select Case Err

        Case 24888
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145342)
        
    End Select
    
    Exit Function
    
End Function

Function CategoriaClienteItem_Le(objCategoriaClienteItem As ClassCategoriaClienteItem) As Long
'Lê na tabela de CategoriaClienteItem a Categoria e o Item de uma deteminada Categoria de Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaClienteItem_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTE_LE_ITEM", X)
    If lErro <> SUCESSO Then Error 24889
    
    
    CategoriaClienteItem_Le = X.CategoriaClienteItem_Le(objCategoriaClienteItem)
    
    Exit Function
    
Erro_CategoriaClienteItem_Le:

    CategoriaClienteItem_Le = Err
    
    Select Case Err

        Case 24889
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145343)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Le_Categorias(objFilialCliente As ClassFilialCliente, colCategoriaItem As Collection) As Long
'Lê na tabela de FilialClienteCategorias, todas as categorias da Filial de Cliente, guardando na coleção colCategoriaItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Le_Categorias
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALCLIENTE_LE_CATEGS", X)
    If lErro <> SUCESSO Then Error 27635
    
    
    
    FilialCliente_Le_Categorias = X.FilialCliente_Le_Categorias(objFilialCliente, colCategoriaItem)
    
    Exit Function
    
Erro_FilialCliente_Le_Categorias:

    FilialCliente_Le_Categorias = Err
    
    Select Case Err

        Case 27635
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145344)
        
    End Select
    
    Exit Function
    
End Function

Function TipoDocInfo_Le_Codigo(objTipoDocInfo As ClassTipoDocInfo) As Long
'objTipoDocInfo.iCodigo tem que estar preenchido
'Preenche objTipoDocInfo a partir do bd p/registro identificado por objTipoDocInfo.iCodigo
'Retorna 31415 se nao encontrar

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Le_Codigo
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPODOCINFO_LE_CODIGO", X)
    If lErro <> SUCESSO Then Error 29677
    
    
    TipoDocInfo_Le_Codigo = X.TipoDocInfo_Le_Codigo(objTipoDocInfo)
    
    Exit Function
    
Erro_TipoDocInfo_Le_Codigo:

    TipoDocInfo_Le_Codigo = Err
    
    Select Case Err

        Case 29677
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145345)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDocInfo_Le(colTipoDocInfo As colTipoDocInfo) As Long
'Lê a Sigla e a Descrição de todos os registros de TiposDocInfo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDocInfo_Le
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSDOCINFO_LE", X)
    If lErro <> SUCESSO Then Error 29678
    
    
    TiposDocInfo_Le = X.TiposDocInfo_Le(colTipoDocInfo)
    
    Exit Function
    
Erro_TiposDocInfo_Le:

    TiposDocInfo_Le = Err
    
    Select Case Err

        Case 29678
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145346)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le(objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal com o Número Interno passado em objNFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE", X)
    If lErro <> SUCESSO Then gError 29679
    
    NFiscal_Le = X.NFiscal_Le(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Le:

    NFiscal_Le = gErr
    
    Select Case gErr

        Case 29679
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 145347)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_Num_Forn(objNFiscal As ClassNFiscal, ByVal iTipo As Integer) As Long
'Lê a Nota Fiscal a partir de TipoNFiscal, Fornecedor, Filial, Serie e NFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Num_Forn
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE_NUM_FORN", X)
    If lErro <> SUCESSO Then Error 29680
    
    
    NFiscal_Le_Num_Forn = X.NFiscal_Le_Num_Forn(objNFiscal, iTipo)
    
    Exit Function
    
Erro_NFiscal_Le_Num_Forn:

    NFiscal_Le_Num_Forn = Err
    
    Select Case Err

        Case 29680
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145348)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_Num_Forn1(objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal a partir de TipoNFiscal, Fornecedor, Filial, Serie e NFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Num_Forn1
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE_NUM_FORN1", X)
    If lErro <> SUCESSO Then Error 55886
    
    
    NFiscal_Le_Num_Forn1 = X.NFiscal_Le_Num_Forn1(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Le_Num_Forn1:

    NFiscal_Le_Num_Forn1 = Err
    
    Select Case Err

        Case 55886
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145349)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalItem_Automatico(lNumIntNFItem As Long) As Long
'Retorna o número do próximo NumIntNFItem disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalItem_Automatico
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCALITEM_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 16812
    
    
    NFiscalItem_Automatico = X.NFiscalItem_Automatico(lNumIntNFItem)
    
    Exit Function
    
Erro_NFiscalItem_Automatico:

    NFiscalItem_Automatico = Err
    
    Select Case Err

        Case 16812
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145350)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Automatico(lNumIntNF As Long) As Long
'Retorna o número do próximo NumIntNF disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Automatico
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 16813
    
    
    NFiscal_Automatico = X.NFiscal_Automatico(lNumIntNF)
    
    Exit Function
    
Erro_NFiscal_Automatico:

    NFiscal_Automatico = Err
    
    Select Case Err

        Case 16813
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145351)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDocInfo_Le_NomeTelaComplemetarTipo(objTipoDocInfo As ClassTipoDocInfo) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDocInfo_Le_NomeTelaComplemetarTipo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPOSDOCINFO_LE_NOMETELACOMPTIPO", X)
    If lErro <> SUCESSO Then Error 19699
    
    
    TiposDocInfo_Le_NomeTelaComplemetarTipo = X.TiposDocInfo_Le_NomeTelaComplemetarTipo(objTipoDocInfo)

    Exit Function
    
Erro_TiposDocInfo_Le_NomeTelaComplemetarTipo:
    
    TiposDocInfo_Le_NomeTelaComplemetarTipo = Err
    
    Select Case Err

        Case 19699
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145352)
        
    End Select
    
    Exit Function

End Function

Function CategoriaClienteItem_Lock(sCategoria As String, sItemCategoria As String, lComando As Long) As Long
'faz lock shared de uma CategoriaClienteItem para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaClienteItem_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CATEGORIACLIENTEITEM_LOCK", X)
    If lErro <> SUCESSO Then Error 19754
    
    
    CategoriaClienteItem_Lock = X.CategoriaClienteItem_Lock(sCategoria, sItemCategoria, lComando)

    Exit Function
    
Erro_CategoriaClienteItem_Lock:
    
    CategoriaClienteItem_Lock = Err
    
    Select Case Err

        Case 19754
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145353)
        
    End Select
    
    Exit Function

End Function

Function TipoDocInfo_Lock_Sigla(sSiglaMovto As String, lComando As Long) As Long
'faz lock shared de um tipodocinfo pela sigla para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Lock_Sigla
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODOCINFO_LOCK_SIGLA", X)
    If lErro <> SUCESSO Then Error 19755
    
    
    TipoDocInfo_Lock_Sigla = X.TipoDocInfo_Lock_Sigla(sSiglaMovto, lComando)

    Exit Function
    
Erro_TipoDocInfo_Lock_Sigla:
    
    TipoDocInfo_Lock_Sigla = Err
    
    Select Case Err

        Case 19755
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145354)
        
    End Select
    
    Exit Function

End Function

Function NaturezaOperacao_Lock(sNaturezaOperacao As String, lComando As Long) As Long
'faz lock shared de uma natureza de operacao para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NaturezaOperacao_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_NATUREZAOPERACAO_LOCK", X)
    If lErro <> SUCESSO Then Error 19753
    
    
    NaturezaOperacao_Lock = X.NaturezaOperacao_Lock(sNaturezaOperacao, lComando)

    Exit Function
    
Erro_NaturezaOperacao_Lock:
    
    NaturezaOperacao_Lock = Err
    
    Select Case Err

        Case 19753
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145355)
        
    End Select
    
    Exit Function

End Function

Function NaturezaOperacao_Exclui(sCodigo As String) As Long
'Exclui a Natureza de Operação do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NaturezaOperacao_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NATUREZAOPERACAO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24518
    
    
    NaturezaOperacao_Exclui = X.NaturezaOperacao_Exclui(sCodigo)

    Exit Function
    
Erro_NaturezaOperacao_Exclui:
    
    NaturezaOperacao_Exclui = Err
    
    Select Case Err

        Case 24518
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145356)
        
    End Select
    
    Exit Function

End Function

Function NaturezaOperacao_Grava(objNaturezaOperacao As ClassNaturezaOp) As Long
'Grava a Natureza de Operação no banco de dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NaturezaOperacao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NATUREZAOPERACAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 24519
    
    
    NaturezaOperacao_Grava = X.NaturezaOperacao_Grava(objNaturezaOperacao)

    Exit Function
    
Erro_NaturezaOperacao_Grava:
    
    NaturezaOperacao_Grava = Err
    
    Select Case Err

        Case 24519
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145357)
        
    End Select
    
    Exit Function

End Function


Function NaturezaOperacao_Le(objNaturezaOperacao As ClassNaturezaOp) As Long
'Lê a natureza da operação a partir da chave fornecida em objNaturezaOperacao. Retorna os dados em objNaturezaOperacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NaturezaOperacao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NATUREZAOPERACAO_LE", X)
    If lErro <> SUCESSO Then Error 24520
    
    
    NaturezaOperacao_Le = X.NaturezaOperacao_Le(objNaturezaOperacao)

    Exit Function
    
Erro_NaturezaOperacao_Le:
    
    NaturezaOperacao_Le = Err
    
    Select Case Err

        Case 24520
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145358)
        
    End Select
    
    Exit Function

End Function

Function NaturezaOperacao_Le_Todos(colNaturezaOperacao As Collection)
'Lê todos as naturezas de operação contidos na tabela de NaturezaOp e coloca-os em colNaturezaOperacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NaturezaOperacao_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NATUREZAOPERACAO_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 24521
    
    
    NaturezaOperacao_Le_Todos = X.NaturezaOperacao_Le_Todos(colNaturezaOperacao)

    Exit Function
    
Erro_NaturezaOperacao_Le_Todos:
    
    NaturezaOperacao_Le_Todos = Err
    
    Select Case Err

        Case 24521
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145359)
        
    End Select
    
    Exit Function

End Function

Function TituloReceber_Grava_Fatura(objTituloReceber As ClassTituloReceber, colNFiscalMarcado As Collection, colParcelaReceber As colParcelaReceber, colcolComissao As colcolComissao, colcolDesconto As colcolDesconto, objContabil As ClassContabil) As Long
'Grava o titulo referente a uma fatura

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Grava_Fatura
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_GRAVA_FATURA", X)
    If lErro <> SUCESSO Then Error 36756
    
    
    TituloReceber_Grava_Fatura = X.TituloReceber_Grava_Fatura(objTituloReceber, colNFiscalMarcado, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)

    Exit Function
    
Erro_TituloReceber_Grava_Fatura:
    
    TituloReceber_Grava_Fatura = Err
    
    Select Case Err

        Case 36756
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145360)
        
    End Select
    
    Exit Function

End Function

Function PedidoVenda_Le_Cliente(objCliente As ClassCliente, lPedido As Long, iFilial As Integer) As Long
'Lê nas tabelas de PedidosVenda e PedidosVendaBaixados o Cliente à partir da chave(lCodigo) fornecida em objCliente.
'Retorna sucesso se achar algum Cliente relacionado com Pedido de Venda.
'Retorna o código do Pedido em lPedido e o da FilialEmpresa em iFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Le_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PEDIDOVENDA_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 19907
    
    
    PedidoVenda_Le_Cliente = X.PedidoVenda_Le_Cliente(objCliente, lPedido, iFilial)

    Exit Function
    
Erro_PedidoVenda_Le_Cliente:
    
    PedidoVenda_Le_Cliente = Err
    
    Select Case Err

        Case 19907
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145361)
        
    End Select
    
    Exit Function

End Function

Function ChequePre_Le_Cliente(objCliente As ClassCliente, lNumIntCheque As Long) As Long
'Lê na tabela de ChequePre, Cliente à partir da chave(lCodigo) fornecida em objCliente.
'Retorna sucesso se achar algum Cliente relacionado à Cheque Pré.
'Em lNumIntCheque retorna o número interno do cheque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Le_Cliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CHEQUEPRE_LE_CLIENTE", X)
    If lErro <> SUCESSO Then Error 19908
    
    
    ChequePre_Le_Cliente = X.ChequePre_Le_Cliente(objCliente, lNumIntCheque)

    Exit Function
    
Erro_ChequePre_Le_Cliente:
    
    ChequePre_Le_Cliente = Err
    
    Select Case Err

        Case 19908
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145362)
        
    End Select
    
    Exit Function

End Function

Function CondicaoPagto_Grava(objCondicaoPagto As ClassCondicaoPagto) As Long
'Atualiza ou insere um novo registro na tabela CondicoesPagto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicaoPagto_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CONDICAOPAGTO_GRAVA", X)
    If lErro <> SUCESSO Then Error 19909
    
    
    CondicaoPagto_Grava = X.CondicaoPagto_Grava(objCondicaoPagto)

    Exit Function
    
Erro_CondicaoPagto_Grava:
    
    CondicaoPagto_Grava = Err
    
    Select Case Err

        Case 19909
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145363)
        
    End Select
    
    Exit Function

End Function

Function CondicaoPagto_Exclui(objCondicaoPagto As ClassCondicaoPagto) As Long
'Exclui CondicaoPagto do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicaoPagto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CONDICAOPAGTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19910
    
    
    CondicaoPagto_Exclui = X.CondicaoPagto_Exclui(objCondicaoPagto)

    Exit Function
    
Erro_CondicaoPagto_Exclui:
    
    CondicaoPagto_Exclui = Err
    
    Select Case Err

        Case 19910
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145364)
        
    End Select
    
    Exit Function

End Function

Function CondicaoPagto_Automatico(iCodigo As Integer) As Long
'Retorna o código da próxima CondicaoPagto disponível

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicaoPagto_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CONDICAOPAGTO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19911
    
    
    CondicaoPagto_Automatico = X.CondicaoPagto_Automatico(iCodigo)

    Exit Function
    
Erro_CondicaoPagto_Automatico:
    
    CondicaoPagto_Automatico = Err
    
    Select Case Err

        Case 19911
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145365)
        
    End Select
    
    Exit Function

End Function

Function PedidosDeVenda_Le_CondicaoPagto(objCondicaoPagto As ClassCondicaoPagto, lPedido As Long, iFilial As Integer) As Long
'Verifica se CondicaoPagto está vinculada à Pedidos de Venda ou à Pedidos de Venda Baixados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosDeVenda_Le_CondicaoPagto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PEDIDOSDEVENDA_LE_CONDICAOPAGTO", X)
    If lErro <> SUCESSO Then Error 19912
    
    
    PedidosDeVenda_Le_CondicaoPagto = X.PedidosDeVenda_Le_CondicaoPagto(objCondicaoPagto, lPedido, iFilial)

    Exit Function
    
Erro_PedidosDeVenda_Le_CondicaoPagto:
    
    PedidosDeVenda_Le_CondicaoPagto = Err
    
    Select Case Err

        Case 19912
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145366)
        
    End Select
    
    Exit Function

End Function

Function PedidoVenda_Le_FilialCliente(objFilialCliente As ClassFilialCliente, lPedido As Long, iFilial As Integer) As Long
'Lê a Filial do Cliente à partir da chave(lCodigo) fornecida em objFilialCliente.
'Retorna sucesso se achar alguma Filial do Cliente relacionada com Pedido de Venda.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Le_FilialCliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PEDIDOVENDA_LE_FILIALCLIENTE", X)
    If lErro <> SUCESSO Then Error 19917
    
    
    PedidoVenda_Le_FilialCliente = X.PedidoVenda_Le_FilialCliente(objFilialCliente, lPedido, iFilial)

    Exit Function
    
Erro_PedidoVenda_Le_FilialCliente:
    
    PedidoVenda_Le_FilialCliente = Err
    
    Select Case Err

        Case 19917
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145367)
        
    End Select
    
    Exit Function

End Function

Function TipoVendedor_Le(objTipoVendedor As ClassTipoVendedor) As Long
'Lê os dados dos Tipos de Vendedor, cujo código foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoVendedor_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOVENDEDOR_LE", X)
    If lErro <> SUCESSO Then Error 19918
    
    
    TipoVendedor_Le = X.TipoVendedor_Le(objTipoVendedor)

    Exit Function
    
Erro_TipoVendedor_Le:
    
    TipoVendedor_Le = Err
    
    Select Case Err

        Case 19918
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145368)
        
    End Select
    
    Exit Function

End Function

Function Vendedor_Le_Tipo(objTipoVendedor As ClassTipoVendedor, iCodigo) As Long
'Tenta ler vendedor relacionado a Tipo de Vendedor passado como parâmetro
'Se existir devolve código do vendedor em iCodigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Le_Tipo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_LE_TIPO", X)
    If lErro <> SUCESSO Then Error 19919
    
    
    Vendedor_Le_Tipo = X.Vendedor_Le_Tipo(objTipoVendedor, iCodigo)

    Exit Function
    
Erro_Vendedor_Le_Tipo:
    
    Vendedor_Le_Tipo = Err
    
    Select Case Err

        Case 19919
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145369)
        
    End Select
    
    Exit Function

End Function

Function TipoVendedor_Grava(objTipoVendedor As ClassTipoVendedor) As Long
'Atualiza ou insere um novo registro na tabela TiposDeVendedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoVendedor_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOVENDEDOR_GRAVA", X)
    If lErro <> SUCESSO Then Error 19920
    
    
    TipoVendedor_Grava = X.TipoVendedor_Grava(objTipoVendedor)

    Exit Function
    
Erro_TipoVendedor_Grava:
    
    TipoVendedor_Grava = Err
    
    Select Case Err

        Case 19920
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145370)
        
    End Select
    
    Exit Function

End Function

Function TipoVendedor_Automatico(iCodigo As Integer) As Long
'Retorna o código do próximo Tipo de Vendedor disponível

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoVendedor_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOVENDEDOR_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19921
    
    
    TipoVendedor_Automatico = X.TipoVendedor_Automatico(iCodigo)

    Exit Function
    
Erro_TipoVendedor_Automatico:
    
    TipoVendedor_Automatico = Err
    
    Select Case Err

        Case 19921
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145371)
        
    End Select
    
    Exit Function

End Function

Function TipoVendedor_Exclui(objTipoVendedor As ClassTipoVendedor) As Long
'exclui Tipo de Vendedor do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoVendedor_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOVENDEDOR_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19922
    
    
    TipoVendedor_Exclui = X.TipoVendedor_Exclui(objTipoVendedor)

    Exit Function
    
Erro_TipoVendedor_Exclui:
    
    TipoVendedor_Exclui = Err
    
    Select Case Err

        Case 19922
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145372)
        
    End Select
    
    Exit Function

End Function

Function Vendedor_Grava(objVendedor As ClassVendedor, objEndereco As ClassEndereco) As Long
'Atualiza ou insere um novo registro nas tabelas Vendedores

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_GRAVA", X)
    If lErro <> SUCESSO Then Error 19923
    
    
    Vendedor_Grava = X.Vendedor_Grava(objVendedor, objEndereco)

    Exit Function
    
Erro_Vendedor_Grava:
    
    Vendedor_Grava = Err
    
    Select Case Err

        Case 19923
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145373)
        
    End Select
    
    Exit Function

End Function

Function TipoVendedor_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'faz lock shared no Tipo Vendedor passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoVendedor_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOVENDEDOR_LOCK", X)
    If lErro <> SUCESSO Then Error 19924
    
    
    TipoVendedor_Lock = X.TipoVendedor_Lock(lComando, iCodigo)

    Exit Function
    
Erro_TipoVendedor_Lock:
    
    TipoVendedor_Lock = Err
    
    Select Case Err

        Case 19924
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145374)
        
    End Select
    
    Exit Function

End Function

Function RegiaoVenda_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'faz lock shared na Regiao de Venda passada como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegiaoVenda_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_REGIAOVENDA_LOCK", X)
    If lErro <> SUCESSO Then Error 19925
    
    
    RegiaoVenda_Lock = X.RegiaoVenda_Lock(lComando, iCodigo)

    Exit Function
    
Erro_RegiaoVenda_Lock:
    
    RegiaoVenda_Lock = Err
    
    Select Case Err

        Case 19925
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145375)
        
    End Select
    
    Exit Function

End Function

Function Vendedor_Exclui(objVendedor As ClassVendedor) As Long
'Exclui Vendedor do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19926
    
    
    Vendedor_Exclui = X.Vendedor_Exclui(objVendedor)

    Exit Function
    
Erro_Vendedor_Exclui:
    
    Vendedor_Exclui = Err
    
    Select Case Err

        Case 19926
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145376)
        
    End Select
    
    Exit Function

End Function

Function Vendedor_Automatico(sCodigo As String) As Long
'retorna o código do próximo Vendedor disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19927
    
    
    Vendedor_Automatico = X.Vendedor_Automatico(sCodigo)

    Exit Function
    
Erro_Vendedor_Automatico:
    
    Vendedor_Automatico = Err
    
    Select Case Err

        Case 19927
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145377)
        
    End Select
    
    Exit Function

End Function

Function Vendedor_Objetos_Relacionados_Testa(objVendedor As ClassVendedor) As Long
'Lê os objetos relacionados com Vendedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Objetos_Relacionados_Testa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_OBJETOS_RELACIONADOS_TEST", X)
    If lErro <> SUCESSO Then Error 19928
    
    
    Vendedor_Objetos_Relacionados_Testa = X.Vendedor_Objetos_Relacionados_Testa(objVendedor)

    Exit Function
    
Erro_Vendedor_Objetos_Relacionados_Testa:
    
    Vendedor_Objetos_Relacionados_Testa = Err
    
    Select Case Err

        Case 19928
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145378)
        
    End Select
    
    Exit Function

End Function

Function Comissao_Le_Vendedor(objVendedor As ClassVendedor, lNumIntCom As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissao_Le_Vendedor
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSAO_LE_VENDEDOR", X)
    If lErro <> SUCESSO Then Error 19929
    
    
    Comissao_Le_Vendedor = X.Comissao_Le_Vendedor(objVendedor, lNumIntCom)

    Exit Function
    
Erro_Comissao_Le_Vendedor:
    
    Comissao_Le_Vendedor = Err
    
    Select Case Err

        Case 19929
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145379)
        
    End Select
    
    Exit Function

End Function

Function ChequePre_Le(objChequePre As ClassChequePre) As Long
'Le o ChequePre a partir da chave(lNumIntCheque) fornecida em objChequePre. Devolve os dados em objChequePre

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CHEQUEPRE_LE", X)
    If lErro <> SUCESSO Then Error 27012
    
    
    ChequePre_Le = X.ChequePre_Le(objChequePre)
    
    Exit Function
    
Erro_ChequePre_Le:

    ChequePre_Le = Err
    
    Select Case Err

        Case 27012
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145380)
        
    End Select
    
    Exit Function
    
End Function

Function CondicaoPagto_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transacao
'faz lock shared na Condicao de Pagamento passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicaoPagto_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CONDICAOPAGTO_LOCK", X)
    If lErro <> SUCESSO Then Error 19906
    
    
    CondicaoPagto_Lock = X.CondicaoPagto_Lock(lComando, iCodigo)
    
    Exit Function
    
Erro_CondicaoPagto_Lock:

    CondicaoPagto_Lock = Err
    
    Select Case Err

        Case 19906
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145381)
        
    End Select
    
    Exit Function
    
End Function
'Janaina - retirei colchequepre
Function TituloReceber_Grava(objTituloReceber As ClassTituloReceber, colComissaoEmissao As colComissao, colParcelaReceber As colParcelaReceber, colcolComissao As colcolComissao, colcolDesconto As colcolDesconto, objContabil As ClassContabil) As Long
'Grava um Título à Receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_GRAVA", X)
    If lErro <> SUCESSO Then Error 19934
    
    
    TituloReceber_Grava = X.TituloReceber_Grava(objTituloReceber, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
    
    Exit Function
    
Erro_TituloReceber_Grava:

    TituloReceber_Grava = Err
    
    Select Case Err

        Case 19934
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145382)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Lock(lCodCliente As Long, iCodFilial As Integer) As Long
'Faz lock da filial do cliente passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALCLIENTE_LOCK", X)
    If lErro <> SUCESSO Then Error 19935
    
    
    FilialCliente_Lock = X.FilialCliente_Lock(lCodCliente, iCodFilial)
    
    Exit Function
    
Erro_FilialCliente_Lock:

    FilialCliente_Lock = Err
    
    Select Case Err

        Case 19935
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145383)
        
    End Select
    
    Exit Function
    
End Function

Function Comissoes_Exclui(iTipoTitulo As Integer, lNumIntDoc As Long, iInclusiveBaixadas As Integer) As Long
'Exclui as comissões do Documento com o Tipo e o Número Interno passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19936
    
    
    Comissoes_Exclui = X.Comissoes_Exclui(iTipoTitulo, lNumIntDoc, iInclusiveBaixadas)
    
    Exit Function
    
Erro_Comissoes_Exclui:

    Comissoes_Exclui = Err
    
    Select Case Err

        Case 19936
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145384)
        
    End Select
    
    Exit Function
    
End Function

Function ChequesPre_Altera(colchequepre As colchequepre, lNumIntTitulo As Long) As Long
'Altera os ChequesPre das Parcelas a Receber que não estão baixadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequesPre_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CHEQUESPRE_ALTERA", X)
    If lErro <> SUCESSO Then Error 19937
    
    
    ChequesPre_Altera = X.ChequesPre_Altera(colchequepre, lNumIntTitulo)
    
    Exit Function
    
Erro_ChequesPre_Altera:

    ChequesPre_Altera = Err
    
    Select Case Err

        Case 19937
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145385)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesParcelas_Altera(colParcelasRec As colParcelaReceber, colcolComissao As colcolComissao) As Long
'Altera as comissões das Parcelas a Receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesParcelas_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOESPARCELAS_ALTERA", X)
    If lErro <> SUCESSO Then Error 19938
    
    
    ComissoesParcelas_Altera = X.ComissoesParcelas_Altera(colParcelasRec, colcolComissao)
    
    Exit Function
    
Erro_ComissoesParcelas_Altera:

    ComissoesParcelas_Altera = Err
    
    Select Case Err

        Case 19938
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145386)
        
    End Select
    
    Exit Function
    
End Function


Function ParcelasReceber_Grava(colParcelaReceber As colParcelaReceber, colchequepre As colchequepre, colcolComissao As colcolComissao) As Long
'Grava as Parcelas de um Titulo passadas por parâmetro
'CHAMADA EM TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasReceber_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASRECEBER_GRAVA", X)
    If lErro <> SUCESSO Then Error 19940
    
    
    ParcelasReceber_Grava = X.ParcelasReceber_Grava(colParcelaReceber, colchequepre, colcolComissao)
    
    Exit Function
    
Erro_ParcelasReceber_Grava:

    ParcelasReceber_Grava = Err
    
    Select Case Err

        Case 19940
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145387)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Exclui(objTituloReceber As ClassTituloReceber, objContabil As ClassContabil) As Long
'Exclui o Título à Receber do Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19941
    
    
    TituloReceber_Exclui = X.TituloReceber_Exclui(objTituloReceber, objContabil)
    
    Exit Function
    
Erro_TituloReceber_Exclui:

    TituloReceber_Exclui = Err
    
    Select Case Err

        Case 19941
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145388)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasReceber_Exclui(objTituloReceber As ClassTituloReceber)
'Exclui as parcelas relacionadas ao Título passado por parâmetro
'FUNÇÃO CHAMADA DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasReceber_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASRECEBER_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19942
    
    
    ParcelasReceber_Exclui = X.ParcelasReceber_Exclui(objTituloReceber)
    
    Exit Function
    
Erro_ParcelasReceber_Exclui:

    ParcelasReceber_Exclui = Err
    
    Select Case Err

        Case 19942
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145389)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Insere(objTituloReceber As ClassTituloReceber, colParcelaReceber As colParcelaReceber, colComissaoEmissao As colComissao, colcolComissao As colcolComissao, colchequepre As colchequepre) As Long
'Insere o Título à Receber na Tabela de Titulos a Receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Insere
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_INSERE", X)
    If lErro <> SUCESSO Then Error 19943
    
    
    TituloReceber_Insere = X.TituloReceber_Insere(objTituloReceber, colParcelaReceber, colComissaoEmissao, colcolComissao, colchequepre)
    
    Exit Function
    
Erro_TituloReceber_Insere:

    TituloReceber_Insere = Err
    
    Select Case Err

        Case 19943
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145390)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Le_Numero(objTituloReceber As ClassTituloReceber) As Long
'Lê o Título com o SiglaDocumento, Numtitulo e Data de Emissao passados em objtituloReceber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_LE_NUMERO", X)
    If lErro <> SUCESSO Then Error 19944
    
    
    TituloReceber_Le_Numero = X.TituloReceber_Le_Numero(objTituloReceber)
    
    Exit Function
    
Erro_TituloReceber_Le_Numero:

    TituloReceber_Le_Numero = Err
    
    Select Case Err

        Case 19944
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145391)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceberBaixado_Le_Numero(objTituloReceber As ClassTituloReceber) As Long
'Lê na Tabela de Titulos Baixados o Titulo com a SiglaDocumento, NumTitulo e Data de Emissao passados em objTituloReceber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceberBaixado_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBERBAIXADO_LE_NUMERO", X)
    If lErro <> SUCESSO Then Error 19946
    
    
    TituloReceberBaixado_Le_Numero = X.TituloReceberBaixado_Le_Numero(objTituloReceber)
    
    Exit Function
    
Erro_TituloReceberBaixado_Le_Numero:

    TituloReceberBaixado_Le_Numero = Err
    
    Select Case Err

        Case 19946
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145392)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDocumento_Le_TituloRec(colSigla As Collection) As Long
'Le as Siglas de Tipos de Documento que tem campo EmTituloRec = 1

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDocumento_Le_TituloRec
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSDOCUMENTO_LE_TITULOREC", X)
    If lErro <> SUCESSO Then Error 19948
    
    
    TiposDocumento_Le_TituloRec = X.TiposDocumento_Le_TituloRec(colSigla)
    
    Exit Function
    
Erro_TiposDocumento_Le_TituloRec:

    TiposDocumento_Le_TituloRec = Err
    
    Select Case Err

        Case 19948
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145393)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Le(objTituloReceber As ClassTituloReceber, Optional iExcluido As Integer = 0) As Long
'Lê Titulo a Receber com o Número Interno passado como parâmetro
'No Parametro Optional: Se iExcluido = 1 então entra os Titulos Excluidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Le
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_LE", X)
    If lErro <> SUCESSO Then Error 19950
    
    TituloReceber_Le = X.TituloReceber_Le(objTituloReceber, iExcluido)
    
    Exit Function
    
Erro_TituloReceber_Le:

    TituloReceber_Le = Err
    
    Select Case Err

        Case 19950
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145394)
        
    End Select
    
    Exit Function
    
End Function

Function DescontoPadrao_Le(colDescontoPadrao As Collection) As Long
'Le os atributos do Desconto Padrão em CPRConfig

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DescontoPadrao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_DESCONTOPADRAO_LE", X)
    If lErro <> SUCESSO Then Error 45828
    
    
    DescontoPadrao_Le = X.DescontoPadrao_Le(colDescontoPadrao)
    
    Exit Function
    
Erro_DescontoPadrao_Le:

    DescontoPadrao_Le = Err
    
    Select Case Err

        Case 45828
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145395)
        
    End Select
    
    Exit Function
    
End Function

Function Comissoes_Le(lNumIntDoc As Long, colComissao As colComissao, iTipoTitulo As Integer) As Long
'Le as Comissões do Documento com o Tipo e o Número Interno passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_LE", X)
    If lErro <> SUCESSO Then Error 19952
    
    
    Comissoes_Le = X.Comissoes_Le(lNumIntDoc, colComissao, iTipoTitulo)
    
    Exit Function
    
Erro_Comissoes_Le:

    Comissoes_Le = Err
    
    Select Case Err

        Case 19952
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145396)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasReceber_Le(objTituloReceber As ClassTituloReceber, colParcelasRec As colParcelaReceber) As Long
'Lê as Parcelas a Receber do Titulo cujo NumIntDoc foi passado em objTituloReceber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasReceber_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASRECEBER_LE", X)
    If lErro <> SUCESSO Then Error 19953
    
    
    ParcelasReceber_Le = X.ParcelasReceber_Le(objTituloReceber, colParcelasRec)
    
    Exit Function
    
Erro_ParcelasReceber_Le:

    ParcelasReceber_Le = Err
    
    Select Case Err

        Case 19953
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145397)
        
    End Select
    
    Exit Function
    
End Function

Function Comissoes_Altera(colComissao As colComissao, iTipoTitulo As Integer, lNumIntDoc As Long) As Long
'Altera no BD as comissoes passadas em colComissao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_ALTERA", X)
    If lErro <> SUCESSO Then Error 19954
    
    
    Comissoes_Altera = X.Comissoes_Altera(colComissao, iTipoTitulo, lNumIntDoc)
    
    Exit Function
    
Erro_Comissoes_Altera:

    Comissoes_Altera = Err
    
    Select Case Err

        Case 19954
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145398)
        
    End Select
    
    Exit Function
    
End Function



Function TituloReceber_Altera(colParcelaReceber As colParcelaReceber, colParcelaReceberBD As colParcelaReceber, colcolDesconto As colcolDesconto, colcolDescontoBD As colcolDesconto, colcolComissao As colcolComissao, colcolComissaoBD As colcolComissao, colchequepre As colchequepre, colChequePreBD As colchequepre) As Long
'Altera os dados do Título à Receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_ALTERA", X)
    If lErro <> SUCESSO Then Error 19957
    
    
    TituloReceber_Altera = X.TituloReceber_Altera(colParcelaReceber, colParcelaReceberBD, colcolDesconto, colcolDescontoBD, colcolComissao, colcolComissaoBD, colchequepre, colChequePreBD)
    
    Exit Function
    
Erro_TituloReceber_Altera:

    TituloReceber_Altera = Err
    
    Select Case Err

        Case 19957
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145399)
        
    End Select
    
    Exit Function
    
End Function

Function Cobrador_Le_Carteiras(objCobrador As ClassCobrador, colCarteirasCobrador As Collection) As Long
'Lê as Carteiras a partir da chave fornecida em iCobrador
'Retorna as Carteiras do Cobrador lidas na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cobrador_Le_Carteiras
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COBRADOR_LE_CARTEIRAS", X)
    If lErro <> SUCESSO Then Error 19958
    
    
    Cobrador_Le_Carteiras = X.Cobrador_Le_Carteiras(objCobrador, colCarteirasCobrador)
    
    Exit Function
    
Erro_Cobrador_Le_Carteiras:

    Cobrador_Le_Carteiras = Err
    
    Select Case Err

        Case 19958
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145400)
        
    End Select
    
    Exit Function
    
End Function

Function CarteiraDeCobranca_Le(objCarteiraCobranca As ClassCarteiraCobranca) As Long
'Lê a Carteira a partir da chave fornecida em objCarteiraCobranca (Código) e retorna os dados em objCarteiraCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraDeCobranca_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRADECOBRANCA_LE", X)
    If lErro <> SUCESSO Then Error 19959
    
    
    CarteiraDeCobranca_Le = X.CarteiraDeCobranca_Le(objCarteiraCobranca)
    
    Exit Function
    
Erro_CarteiraDeCobranca_Le:

    CarteiraDeCobranca_Le = Err
    
    Select Case Err

        Case 19959
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145401)
        
    End Select
    
    Exit Function
    
End Function

'??? retirar do dicdados
'Function Valida_Exclusao_Cobrador(objCobrador As ClassCobrador) As Long
''verifica a presença do cobrador passado nas tabelas
''BorderosCobranca, OcorrenciasRemParcRec, FiliaisClientes e BaixasParcRec
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Valida_Exclusao_Cobrador
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CRFAT_VALIDA_EXCLUSAO_COBRADOR",X)
'    If lErro <> SUCESSO Then Error 19960
'
'
'    Valida_Exclusao_Cobrador = X.Valida_Exclusao_Cobrador(objCobrador)
'
'    Exit Function
'
'Erro_Valida_Exclusao_Cobrador:
'
'    Valida_Exclusao_Cobrador = Err
'
'    Select Case Err
'
'        Case 19960
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145402)
'
'    End Select
'
'    Exit Function
'
'End Function
'
Function CarteiraCobrador_Le(objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Lê a Carteira a partir da chave fornecida em objCarteiraCobrador
'e devolve dados em objCarteiraCobrador

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraCobrador_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRACOBRADOR_LE", X)
    If lErro <> SUCESSO Then Error 19961
    
    
    CarteiraCobrador_Le = X.CarteiraCobrador_Le(objCarteiraCobrador)
    
    Exit Function
    
Erro_CarteiraCobrador_Le:

    CarteiraCobrador_Le = Err
    
    Select Case Err

        Case 19961
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145403)
        
    End Select
    
    Exit Function
    
End Function

Function CarteirasDeCobranca_Le_Todas(colCarteirasCobranca As Collection)
'Lê todos as Carteiras contidas na tabela CarteirasCobranca por ordem de codigo e as coloca na coleçao colCarteirasCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteirasDeCobranca_Le_Todas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRASDECOBRANCA_LE_TODAS", X)
    If lErro <> SUCESSO Then Error 19964
    
    
    CarteirasDeCobranca_Le_Todas = X.CarteirasDeCobranca_Le_Todas(colCarteirasCobranca)
    
    Exit Function
    
Erro_CarteirasDeCobranca_Le_Todas:

    CarteirasDeCobranca_Le_Todas = Err
    
    Select Case Err

        Case 19964
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145404)
        
    End Select
    
    Exit Function
    
End Function


Function CarteiraDeCobranca_Grava(objCarteiraCobranca As ClassCarteiraCobranca) As Long
'Grava a Carteira no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraDeCobranca_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRASDECOBRANCA_GRAVA", X)
    If lErro <> SUCESSO Then Error 19966
    
    
    CarteiraDeCobranca_Grava = X.CarteiraDeCobranca_Grava(objCarteiraCobranca)
    
    Exit Function
    
Erro_CarteiraDeCobranca_Grava:

    CarteiraDeCobranca_Grava = Err
    
    Select Case Err

        Case 19966
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145405)
        
    End Select
    
    Exit Function
    
End Function

Function CarteiraDeCobranca_Exclui(iCodigo As Integer) As Long
'Exclui a Carteira do Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraDeCobranca_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRASDECOBRANCA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19967
    
    
    CarteiraDeCobranca_Exclui = X.CarteiraDeCobranca_Exclui(iCodigo)
    
    Exit Function
    
Erro_CarteiraDeCobranca_Exclui:

    CarteiraDeCobranca_Exclui = Err
    
    Select Case Err

        Case 19967
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145406)
        
    End Select
    
    Exit Function
    
End Function

Function TipoInstrCobranca_Le(objTipoInstrCobr As ClassTipoInstrCobr) As Long
'le os dados do Tipo Instrucao Cobranca, cujo codigo foi passado como parametro em objTipoInstrCobr

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoInstrCobranca_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOINSTRCOBRANCA_LE", X)
    If lErro <> SUCESSO Then Error 18519
    
    
    TipoInstrCobranca_Le = X.TipoInstrCobranca_Le(objTipoInstrCobr)
    
    Exit Function
    
Erro_TipoInstrCobranca_Le:

    TipoInstrCobranca_Le = Err
    
    Select Case Err

        Case 18519
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145407)
        
    End Select
    
    Exit Function
    
End Function

Function TiposCliente_Le_PadraoCobranca(iCodigo As Integer) As Long
'le os tipos de cliente relacionados ao codigo do Padrao de Cobranca passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposCliente_Le_PadraoCobranca
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSCLIENTE_LE_PADRAOCOBRANCA", X)
    If lErro <> SUCESSO Then Error 18520
    
    
    TiposCliente_Le_PadraoCobranca = X.TiposCliente_Le_PadraoCobranca(iCodigo)
    
    Exit Function
    
Erro_TiposCliente_Le_PadraoCobranca:

    TiposCliente_Le_PadraoCobranca = Err
    
    Select Case Err

        Case 18520
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145408)
        
    End Select
    
    Exit Function
    
End Function

Function Clientes_Le_PadraoCobranca(iCodigo As Integer) As Long
'le  clientes relacionados ao codigo do Padrao de Cobranca passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Clientes_Le_PadraoCobranca
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CLIENTES_LE_PADRAOCOBRANCA", X)
    If lErro <> SUCESSO Then Error 18521
    
    
    Clientes_Le_PadraoCobranca = X.Clientes_Le_PadraoCobranca(iCodigo)
    
    Exit Function
    
Erro_Clientes_Le_PadraoCobranca:

    Clientes_Le_PadraoCobranca = Err
    
    Select Case Err

        Case 18521
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145409)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoCobranca_Grava(objPadraoCobranca As ClassPadraoCobranca) As Long
'Atualiza ou insere um novo registro na tabela PadroesCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOCOBRANCA_GRAVA", X)
    If lErro <> SUCESSO Then Error 18522
    
    
    PadraoCobranca_Grava = X.PadraoCobranca_Grava(objPadraoCobranca)
    
    Exit Function
    
Erro_PadraoCobranca_Grava:

    PadraoCobranca_Grava = Err
    
    Select Case Err

        Case 18522
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145410)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoCobranca_Exclui(objPadraoCobranca As ClassPadraoCobranca) As Long
'exclui Padrao Cobranca do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOCOBRANCA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 18523
    
    
    PadraoCobranca_Exclui = X.PadraoCobranca_Exclui(objPadraoCobranca)
    
    Exit Function
    
Erro_PadraoCobranca_Exclui:

    PadraoCobranca_Exclui = Err
    
    Select Case Err

        Case 18523
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145411)
        
    End Select
    
    Exit Function
    
End Function


Function InstrCobranca_Lock(iInstrucao As Integer, lComando As Long) As Long
'faz lock shared de uma Instrucao de Cobranca  para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InstrCobranca_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_INSTRCOBRANCA_LOCK", X)
    If lErro <> SUCESSO Then Error 18525
    
    
    InstrCobranca_Lock = X.InstrCobranca_Lock(iInstrucao, lComando)
    
    Exit Function
    
Erro_InstrCobranca_Lock:

    InstrCobranca_Lock = Err
    
    Select Case Err

        Case 18525
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145412)
        
    End Select
    
    Exit Function
    
End Function

Function Transportadora_Grava(objTransportadora As ClassTransportadora, objEndereco As ClassEndereco) As Long
'Grava A Transportadora no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Transportadora_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSPORTADORA_GRAVA", X)
    If lErro <> SUCESSO Then Error 19990
    
    
    Transportadora_Grava = X.Transportadora_Grava(objTransportadora, objEndereco)
    
    Exit Function
    
Erro_Transportadora_Grava:

    Transportadora_Grava = Err
    
    Select Case Err

        Case 19990
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145413)
        
    End Select
    
    Exit Function
    
End Function

Function Transportadora_Automatico(iCodigo As Integer) As Long
'obtem o proximo numero da Transportadora
'chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Transportadora_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSPORTADORA_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19991
    
    
    Transportadora_Automatico = X.Transportadora_Automatico(iCodigo)
    
    Exit Function
    
Erro_Transportadora_Automatico:

    Transportadora_Automatico = Err
    
    Select Case Err

        Case 19991
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145414)
        
    End Select
    
    Exit Function
    
End Function

Function Parcela_Com_Desconto(ColDesconto As ColDesconto, dValorParcela As Double, dtDataDeposito As Date, dtDataVencimento As Date, dDesconto As Double) As Long
'Retorna o Valor da Parcela com o desconto mais apropriado dos descontos contidos em colDescoltos, já aplicados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Parcela_Com_Desconto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELA_COM_DESCONTO", X)
    If lErro <> SUCESSO Then Error 19992
    
    
    Parcela_Com_Desconto = X.Parcela_Com_Desconto(ColDesconto, dValorParcela, dtDataDeposito, dtDataVencimento, dDesconto)
    
    Exit Function
    
Erro_Parcela_Com_Desconto:

    Parcela_Com_Desconto = Err
    
    Select Case Err

        Case 19992
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145415)
        
    End Select
    
    Exit Function
    
End Function
''
''Function ChequePre_Inserir(objChequePre As ClassChequePre) As Long
'''Rotina responsavel pela insercao de Cheque Pre
'''TEM QUE SER CHAMADA DENTRO DE TRANSACAO
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_ChequePre_Inserir
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CPR_CHEQUEPRE_INSERIR",X)
''    If lErro <> SUCESSO Then Error 27013
''
''
''    ChequePre_Inserir = X.ChequePre_Inserir(objChequePre)
''
''    Exit Function
''
''Erro_ChequePre_Inserir:
''
''    ChequePre_Inserir = Err
''
''    Select Case Err
''
''        Case 27013
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145416)
''
''    End Select
''
''    Exit Function
''
''End Function

Function DebitoRecCli_Gera(objNFiscal As ClassNFiscal, sTipoDocCPR As String, lNumIntDocCPR As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DebitoRecCli_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_DEBITORECCLI_GERA", X)
    If lErro <> SUCESSO Then Error 45684
    
    
    DebitoRecCli_Gera = X.DebitoRecCli_Gera(objNFiscal, sTipoDocCPR, lNumIntDocCPR)
    
    Exit Function
    
Erro_DebitoRecCli_Gera:

    DebitoRecCli_Gera = Err
    
    Select Case Err

        Case 45684
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145417)
        
    End Select
    
    Exit Function
    
End Function


Function DebitoRecCli_Insere(objDebitoRecCli As ClassDebitoRecCli) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DebitoRecCli_Insere
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_DEBITORECCLI_INSERE", X)
    If lErro <> SUCESSO Then Error 45683
    
    
    DebitoRecCli_Insere = X.DebitoRecCli_Insere(objDebitoRecCli)
    
    Exit Function
    
Erro_DebitoRecCli_Insere:

    DebitoRecCli_Insere = Err
    
    Select Case Err

        Case 45683
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145418)
        
    End Select
    
    Exit Function
    
End Function

Function Cobrador_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared do Cobrador passado como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_Cobrador_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_COBRADOR_LOCK", X)
    If lErro <> SUCESSO Then Error 45371
    
    
    Cobrador_Lock = X.Cobrador_Lock(lComando, iCodigo)
    
    Exit Function
    
Erro_Cobrador_Lock:

    Cobrador_Lock = Err
    
    Select Case Err

        Case 45371
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145419)
        
    End Select
    
    Exit Function
    
End Function

Function Mensagem_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared da Mensagem passada como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_Mensagem_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_MENSAGEM_LOCK", X)
    If lErro <> SUCESSO Then Error 45370
    
    
    Mensagem_Lock = X.Mensagem_Lock(lComando, iCodigo)
    
    Exit Function
    
Erro_Mensagem_Lock:

    Mensagem_Lock = Err
    
    Select Case Err

        Case 45370
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145420)
        
    End Select
    
    Exit Function
    
End Function

Function CPRConfig_ObterNumInt(sCodigo As String, lNumInt As Long, Optional iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CPRConfig_ObterNumInt
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CPRCONFIG_OBTERNUMINT", X)
    If lErro <> SUCESSO Then Error 14790
    
    
    CPRConfig_ObterNumInt = X.CPRConfig_ObterNumInt(sCodigo, lNumInt, iQtde)
    
    Exit Function
    
Erro_CPRConfig_ObterNumInt:

    CPRConfig_ObterNumInt = Err
    
    Select Case Err

        Case 14790
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145421)
        
    End Select
    
    Exit Function
    
End Function

Function ChequePre_Automatico(lNumIntDoc) As Long
'Gera o próximo Número Interno de um ChequePre
'CHAMADA EM TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CHEQUEPRE_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19955
    
    
    ChequePre_Automatico = X.ChequePre_Automatico(lNumIntDoc)
    
    Exit Function
    
Erro_ChequePre_Automatico:

    ChequePre_Automatico = Err
    
    Select Case Err

        Case 19955
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145422)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoCobranca_Automatico(sCodigo As String) As Long
'retorna o código do próximo PadraoCobranca disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOCOBRANCA_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 18524
    
    
    PadraoCobranca_Automatico = X.PadraoCobranca_Automatico(sCodigo)
    
    Exit Function
    
Erro_PadraoCobranca_Automatico:

    PadraoCobranca_Automatico = Err
    
    Select Case Err

        Case 18524
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145423)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Automatico(lNumIntDoc) As Long
'Gera o próximo Número Interno de um Título a Receber
'CHAMADA EM TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19939
    
    
    TituloReceber_Automatico = X.TituloReceber_Automatico(lNumIntDoc)
    
    Exit Function
    
Erro_TituloReceber_Automatico:

    TituloReceber_Automatico = Err
    
    Select Case Err

        Case 19939
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145424)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelaReceber_Automatico(lNumIntDoc) As Long
'Gera o próximo Número Interno a ser utilizado por uma Parcela à Receber
'CHAMADA EM TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelaReceber_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELARECEBER_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19956
    
    
    ParcelaReceber_Automatico = X.ParcelaReceber_Automatico(lNumIntDoc)
    
    Exit Function
    
Erro_ParcelaReceber_Automatico:

    ParcelaReceber_Automatico = Err
    
    Select Case Err

        Case 19956
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145425)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_ObterDadosBasicos1(lComando As Long, objCliente As ClassCliente) As Long
'le o cliente a partir da chave(lCodigo) fornecida em objCliente.
'Devolve os dados em objCliente
'nao acessa dados da matriz do cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cliente_ObterDadosBasicos1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CLIENTE_OBTERDADOSBASICOS1", X)
    If lErro <> SUCESSO Then Error 36982
    
    
    Cliente_ObterDadosBasicos1 = X.Cliente_ObterDadosBasicos1(lComando, objCliente)
    
    Exit Function
    
Erro_Cliente_ObterDadosBasicos1:

    Cliente_ObterDadosBasicos1 = Err
    
    Select Case Err

        Case 36982
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145426)
        
    End Select
    
    Exit Function
    
End Function

Function Pais_Lock(iCodigoPais As Integer, lComando As Long) As Long
'Rotina executada dentro de uma transacao
'faz lock shared no Codigo País passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pais_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PAIS_LOCK", X)
    If lErro <> SUCESSO Then Error 45020
    
    
    Pais_Lock = X.Pais_Lock(iCodigoPais, lComando)
    
    Exit Function
    
Erro_Pais_Lock:

    Pais_Lock = Err
    
    Select Case Err

        Case 45020
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145427)
        
    End Select
    
    Exit Function
    
End Function

Function CPR_Gera(objNFiscal As ClassNFiscal, iClasseDocCPR As Integer, lNumIntDocCPR As Long) As Long
'Gera conta a Pagar/Receber vinculada à Nota Fiscal.
'Devolve ClasseDoc e NumIntDoc de Conta a Pagar/Receber gravada.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CPR_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CPR_GERA", X)
    If lErro <> SUCESSO Then Error 45097
    
    
    CPR_Gera = X.CPR_Gera(objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    
    Exit Function
    
Erro_CPR_Gera:

    CPR_Gera = Err
    
    Select Case Err

        Case 45097
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145428)
        
    End Select
    
    Exit Function
    
End Function

Function TituloPagar_Gera(objNFiscal As ClassNFiscal, sTipoDocCPR As String, lNumIntDocCPR As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloPagar_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULOPAGAR_GERA", X)
    If lErro <> SUCESSO Then Error 45098
    
    
    TituloPagar_Gera = X.TituloPagar_Gera(objNFiscal, sTipoDocCPR, lNumIntDocCPR)
    
    Exit Function
    
Erro_TituloPagar_Gera:

    TituloPagar_Gera = Err
    
    Select Case Err

        Case 45098
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145429)
        
    End Select
    
    Exit Function
    
End Function

Function NFPagar_Gera(objNFiscal As ClassNFiscal, sTipoDocCPR As String, lNumIntDocCPR As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFPagar_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFPAGAR_GERA", X)
    If lErro <> SUCESSO Then Error 45099
    
    
    NFPagar_Gera = X.NFPagar_Gera(objNFiscal, sTipoDocCPR, lNumIntDocCPR)
    
    Exit Function
    
Erro_NFPagar_Gera:

    NFPagar_Gera = Err
    
    Select Case Err

        Case 45099
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145430)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Gera(objNFiscal As ClassNFiscal, sTipoDocCPR As String, lNumIntDocCPR As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_GERA", X)
    If lErro <> SUCESSO Then Error 45100
    
    
    TituloReceber_Gera = X.TituloReceber_Gera(objNFiscal, sTipoDocCPR, lNumIntDocCPR)
    
    Exit Function
    
Erro_TituloReceber_Gera:

    TituloReceber_Gera = Err
    
    Select Case Err

        Case 45100
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145431)
        
    End Select
    
    Exit Function
    
End Function

Function NFFatPag_Grava_BD(objTituloPagar As ClassTituloPagar, colParcelaPagar As colParcelaPagar) As Long
'Grava NotaFiscalFatura no BD (nível + baixo)
'CHAMADA EM TRANSAÇÃO por NFFatPag_Grava

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFFatPag_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_NFFATPAG_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 24127
    
    
    NFFatPag_Grava_BD = X.NFFatPag_Grava_BD(objTituloPagar, colParcelaPagar)
    
    Exit Function
    
Erro_NFFatPag_Grava_BD:

    NFFatPag_Grava_BD = Err
    
    Select Case Err

        Case 24127
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145432)
        
    End Select
    
    Exit Function
    
End Function

Function NFPag_Grava_BD(objNFsPag As ClassNFsPag) As Long
'Grava Nota Fiscal a Pagar no BD
'Chamada por NFPag_Grava DENTRO DE TANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFPag_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_NFPAG_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 14944
    
    
    NFPag_Grava_BD = X.NFPag_Grava_BD(objNFsPag)
    
    Exit Function
    
Erro_NFPag_Grava_BD:

    NFPag_Grava_BD = Err
    
    Select Case Err

        Case 14944
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145433)
        
    End Select
    
    Exit Function
    
End Function

Function DataVencto_Real(dtDataVencimento As Date, dtDataVenctoReal As Date) As Long
'Gera a Data de Vencimento Real em cima da Data de Vencimento passada por parâmetro
'AVISO: ESSA ROTINA TEM QUE SER CHAMADA DENTRO DE UMA TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DataVencto_Real
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_DATAVENCTO_REAL", X)
    If lErro <> SUCESSO Then Error 24119
    
    
    DataVencto_Real = X.DataVencto_Real(dtDataVencimento, dtDataVenctoReal)
    
    Exit Function
    
Erro_DataVencto_Real:

    DataVencto_Real = Err
    
    Select Case Err

        Case 24119
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145434)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function Comissoes_Gera(lComando As Long, lComando1 As Long, objNFiscal As ClassNFiscal, colComissoesEmissao As colComissao) As Long
'Gera as Comissões na Emissão com base nas comissões do grid comissão em coloca em colComissoesEmissao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_GERA", X)
    If lErro <> SUCESSO Then Error 44344
    
    
    Comissoes_Gera = X.Comissoes_Gera(lComando, lComando1, objNFiscal, colComissoesEmissao)
    
    Exit Function
    
Erro_Comissoes_Gera:

    Comissoes_Gera = Err
    
    Select Case Err

        Case 44344
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145435)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function Comissao_Automatico1(lComando As Long, lComando1 As Long, lNumIntDoc As Long) As Long
'Gera o próximo Número Interno de uma Comissao
'AVISO: ESSA ROTINA DEVE SER CHAMADA DE DENTRO DE UMA TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissao_Automatico1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSAO_AUTOMATICO1", X)
    If lErro <> SUCESSO Then Error 44345
    
    
    Comissao_Automatico1 = X.Comissao_Automatico1(lComando, lComando1, lNumIntDoc)
    
    Exit Function
    
Erro_Comissao_Automatico1:

    Comissao_Automatico1 = Err
    
    Select Case Err

        Case 44345
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145436)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function Comissoes_Grava1(lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long, colComissao As colComissao) As Long
'Grava as Comissões passadas em colComissao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comissoes_Grava1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOES_GRAVA1", X)
    If lErro <> SUCESSO Then Error 44346
    
    
    Comissoes_Grava1 = X.Comissoes_Grava1(lComando, lComando1, lComando2, lComando3, colComissao)
    
    Exit Function
    
Erro_Comissoes_Grava1:

    Comissoes_Grava1 = Err
    
    Select Case Err

        Case 44346
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145437)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function NFiscal_Automatico1(lComando As Long, lComando1 As Long, lComando2 As Long, lNumIntNF As Long) As Long
'Retorna o número do próximo NumIntNF disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Automatico1
    


    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_AUTOMATICO1", X)
    If lErro <> SUCESSO Then Error 44347
    
    
    NFiscal_Automatico1 = X.NFiscal_Automatico1(lComando, lComando1, lComando2, lNumIntNF)
    
    Exit Function
    
Erro_NFiscal_Automatico1:

    NFiscal_Automatico1 = Err
    
    Select Case Err

        Case 44347
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145438)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function ComissoesNF_Grava(lComando As Long, lComando1 As Long, objNFiscal As ClassNFiscal) As Long
'Insere no Banco de Dados as Comissões da Nota Fiscal passada em objNFiscal
'ATENÇÃO: Essa função deve ser chamada dentro de Transação.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesNF_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOESNF_GRAVA", X)
    If lErro <> SUCESSO Then Error 44349
    
    
    ComissoesNF_Grava = X.ComissoesNF_Grava(lComando, lComando1, objNFiscal)
    
    Exit Function
    
Erro_ComissoesNF_Grava:

    ComissoesNF_Grava = Err
    
    Select Case Err

        Case 44349
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145439)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function FilialCliente_Lock1(lComando As Long, lCodCliente As Long, iCodFilial As Integer) As Long
'Faz lock da filial do cliente passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Lock1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALCLIENTE_LOCK1", X)
    If lErro <> SUCESSO Then Error 44350
    
    
    FilialCliente_Lock1 = X.FilialCliente_Lock1(lComando, lCodCliente, iCodFilial)
    
    Exit Function
    
Erro_FilialCliente_Lock1:

    FilialCliente_Lock1 = Err
    
    Select Case Err

        Case 44350
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145440)
        
    End Select
    
    Exit Function
    
End Function

Function CanalVenda_Lock(lComando As Long, iCanalVenda As Integer) As Long
'Faz lock do Canal de Venda passado como parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CanalVenda_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CANALVENDA_LOCK", X)
    If lErro <> SUCESSO Then Error 44351
    
    
    CanalVenda_Lock = X.CanalVenda_Lock(lComando, iCanalVenda)
    
    Exit Function
    
Erro_CanalVenda_Lock:

    CanalVenda_Lock = Err
    
    Select Case Err

        Case 44351
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145441)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function Serie_Lock(lComando As Long, sSerie As String) As Long
'Faz lock da Série passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_SERIE_LOCK", X)
    If lErro <> SUCESSO Then Error 44352
    
    
    Serie_Lock = X.Serie_Lock(lComando, sSerie)
    
    Exit Function
    
Erro_Serie_Lock:

    Serie_Lock = Err
    
    Select Case Err

        Case 44352
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145442)
        
    End Select
    
    Exit Function
    
End Function


'mario
Function Transportadora_Lock(lComando As Long, iTransportadora As Integer) As Long
'Faz lock da Transportadora passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Transportadora_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSPORTADORA_LOCK", X)
    If lErro <> SUCESSO Then Error 44353
    
    
    Transportadora_Lock = X.Transportadora_Lock(lComando, iTransportadora)
    
    Exit Function
    
Erro_Transportadora_Lock:

    Transportadora_Lock = Err
    
    Select Case Err

        Case 44353
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145443)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function Vendedor_Lock(ByVal lComando As Long, iVendedor As Integer) As Long
'Faz lock do Vendedor passado como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Vendedor_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_VENDEDOR_LOCK", X)
    If lErro <> SUCESSO Then Error 44354
    
    
    Vendedor_Lock = X.Vendedor_Lock(lComando, iVendedor)
    
    Exit Function
    
Erro_Vendedor_Lock:

    Vendedor_Lock = Err
    
    Select Case Err

        Case 44354
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145444)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function CRFATConfig_ObterNumInt(lComandoSel As Long, lComandoMod As Long, sCodigo As String, lNumInt As Long, Optional iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_ObterNumInt
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CRFATCONFIG_OBTERNUMINT", X)
    If lErro <> SUCESSO Then Error 44355
    
    
    
    If IsMissing(iQtde) Then
        CRFATConfig_ObterNumInt = X.CRFATConfig_ObterNumInt(lComandoSel, lComandoMod, sCodigo, lNumInt)
    Else
        CRFATConfig_ObterNumInt = X.CRFATConfig_ObterNumInt(lComandoSel, lComandoMod, sCodigo, lNumInt, iQtde)
    End If
    
    Exit Function
    
Erro_CRFATConfig_ObterNumInt:

    CRFATConfig_ObterNumInt = Err
    
    Select Case Err

        Case 44355
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145445)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function ComissoesNF_Le(objNFiscal As ClassNFiscal) As Long
'Le as Comissões da Nota Fiscal com o Número Interno passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesNF_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOESNF_LE", X)
    If lErro <> SUCESSO Then Error 44360
    
    
    ComissoesNF_Le = X.ComissoesNF_Le(objNFiscal)
    
    Exit Function
    
Erro_ComissoesNF_Le:

    ComissoesNF_Le = Err
    
    Select Case Err

        Case 44360
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145446)
        
    End Select
    
    Exit Function
    
End Function

'mario
Function ParcelasRecNF_Le(objNFiscal As ClassNFiscal) As Long
'Lê as parcelas a receber vinculadas à Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasRecNF_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASRECNF_LE", X)
    If lErro <> SUCESSO Then Error 44370
    
    
    ParcelasRecNF_Le = X.ParcelasRecNF_Le(objNFiscal)
    
    Exit Function
    
Erro_ParcelasRecNF_Le:

    ParcelasRecNF_Le = Err
    
    Select Case Err

        Case 44370
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145447)
        
    End Select
    
    Exit Function
    
End Function

Function CarteiraCobrador_AtualizaSaldo(objCarteiraCobrador As ClassCarteiraCobrador, iQtdeParcelas As Integer, dValorParcelas As Double) As Long
'atualiza a qtde e valor de titulos na carteira de cobranca de um cobrador
'iQtdeParcelas e dValorParcelas devem ser positivos para entradas e negativos para saidas da carteira
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraCobrador_AtualizaSaldo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CARTEIRACOBRADOR_ATUALIZASALDO", X)
    If lErro <> SUCESSO Then Error 45341
    
    
    CarteiraCobrador_AtualizaSaldo = X.CarteiraCobrador_AtualizaSaldo(objCarteiraCobrador, iQtdeParcelas, dValorParcelas)
    
    Exit Function
    
Erro_CarteiraCobrador_AtualizaSaldo:

    CarteiraCobrador_AtualizaSaldo = Err
    
    Select Case Err

        Case 45341
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145448)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Le_SemNumIntDoc(objTituloReceber As ClassTituloReceber) As Long
'Lê dados do Título a partir de FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Le_SemNumIntDoc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_LE_SEMNUMINTDOC", X)
    If lErro <> SUCESSO Then Error 45354
    
    
    TituloReceber_Le_SemNumIntDoc = X.TituloReceber_Le_SemNumIntDoc(objTituloReceber)
    
    Exit Function
    
Erro_TituloReceber_Le_SemNumIntDoc:

    TituloReceber_Le_SemNumIntDoc = Err
    
    Select Case Err

        Case 45354
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145449)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelaReceber_Le_SemNumIntDoc(objParcelaReceber As ClassParcelaReceber) As Long
'Lê Número Interno da Parcela a partir de NumIntTitulo e NumParcela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelaReceber_Le_SemNumIntDoc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELARECEBER_LE_SEMNUMINTDOC", X)
    If lErro <> SUCESSO Then Error 45355
    
    
    ParcelaReceber_Le_SemNumIntDoc = X.ParcelaReceber_Le_SemNumIntDoc(objParcelaReceber)
    
    Exit Function
    
Erro_ParcelaReceber_Le_SemNumIntDoc:

    ParcelaReceber_Le_SemNumIntDoc = Err
    
    Select Case Err

        Case 45355
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145450)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelaReceberBaixada_Le_SemNumIntDoc(objParcelaReceber As ClassParcelaReceber) As Long
'Lê Número Interno da Parcela a partir de NumIntTitulo e NumParcela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelaReceberBaixada_Le_SemNumIntDoc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELARECBAIXADA_LE_SEMNUMINTDOC", X)
    If lErro <> SUCESSO Then Error 45356
    
    
    ParcelaReceberBaixada_Le_SemNumIntDoc = X.ParcelaReceberBaixada_Le_SemNumIntDoc(objParcelaReceber)
    
    Exit Function
    
Erro_ParcelaReceberBaixada_Le_SemNumIntDoc:

    ParcelaReceberBaixada_Le_SemNumIntDoc = Err
    
    Select Case Err

        Case 45356
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145451)
        
    End Select
    
    Exit Function
    
End Function

Function TiposInstrCobranca_Le_Todos(colTiposInstrCobranca As Collection) As Long
'Carrega uma coleção com objetos ClassTipoInstrCobr para todos os tipos da tabela TiposInstrCobranca

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposInstrCobranca_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSINSTRCOBRANCA_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 45357
    
    
    TiposInstrCobranca_Le_Todos = X.TiposInstrCobranca_Le_Todos(colTiposInstrCobranca)
    
    Exit Function
    
Erro_TiposInstrCobranca_Le_Todos:

    TiposInstrCobranca_Le_Todos = Err
    
    Select Case Err

        Case 45357
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145452)
        
    End Select
    
    Exit Function
    
End Function

Function OcorrRemParcRec_Le(objOcorrRemParcRec As ClassOcorrRemParcRec) As Long
'Traz do bd as informações para preencher um objeto da classe ClassOcorrRemParcRec à partir
'da chave fornecida em objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr e objOcorrRemParcRec.iFilialEmpresa que devem estar preenchidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OcorrRemParcRec_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_OCORRREMPARCREC_LE", X)
    If lErro <> SUCESSO Then Error 45358
    
    
    OcorrRemParcRec_Le = X.OcorrRemParcRec_Le(objOcorrRemParcRec)
    
    Exit Function
    
Erro_OcorrRemParcRec_Le:

    OcorrRemParcRec_Le = Err
    
    Select Case Err

        Case 45358
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145453)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDeCliente_Grava(objTipoCliente As ClassTipoCliente) As Long
'Inclui ou altera um Tipo de Cliente
'Tabelas:TiposDeCliente e TipoDeClienteCategoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDeCliente_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSDECLIENTE_GRAVA", X)
    If lErro <> SUCESSO Then Error 45359
    
    
    TiposDeCliente_Grava = X.TiposDeCliente_Grava(objTipoCliente)
    
    Exit Function
    
Erro_TiposDeCliente_Grava:

    TiposDeCliente_Grava = Err
    
    Select Case Err

        Case 45359
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145454)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaDePreco_Lock(iTabelaPreco As Integer) As Long
'Faz lock da Tabela de Preço passada como parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaDePreco_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELADEPRECO_LOCK", X)
    If lErro <> SUCESSO Then Error 45360
    
    
    TabelaDePreco_Lock = X.TabelaDePreco_Lock(iTabelaPreco)
    
    Exit Function
    
Erro_TabelaDePreco_Lock:

    TabelaDePreco_Lock = Err
    
    Select Case Err

        Case 45360
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145455)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoCobranca_Lock(lComando As Long, iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared do Padrão de Cobrança passada como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoCobranca_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOCOBRANCA_LOCK", X)
    If lErro <> SUCESSO Then Error 45363
    
    
    PadraoCobranca_Lock = X.PadraoCobranca_Lock(lComando, iCodigo)
    
    Exit Function
    
Erro_PadraoCobranca_Lock:

    PadraoCobranca_Lock = Err
    
    Select Case Err

        Case 45363
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145456)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDeCliente_Exclui(objTipoCliente As ClassTipoCliente) As Long
'Exclui o TipoCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDeCliente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOSDECLIENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45364
    
    
    TiposDeCliente_Exclui = X.TiposDeCliente_Exclui(objTipoCliente)
    
    Exit Function
    
Erro_TiposDeCliente_Exclui:

    TiposDeCliente_Exclui = Err
    
    Select Case Err

        Case 45364
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145457)
        
    End Select
    
    Exit Function
    
End Function

Function TransfCartCobr_Grava(objTransfCartCobr As ClassTransfCartCobr) As Long
'Atualiza a tabela de ParcelasRec e Insere os dados na tabela de TransfCartCobr
'Tambem atualiza as estatisticas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TransfCartCobr_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSFCARTCOBR_GRAVA", X)
    If lErro <> SUCESSO Then Error 45365
    
    
    TransfCartCobr_Grava = X.TransfCartCobr_Grava(objTransfCartCobr)
    
    Exit Function
    
Erro_TransfCartCobr_Grava:

    TransfCartCobr_Grava = Err
    
    Select Case Err

        Case 45365
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145458)
        
    End Select
    
    Exit Function
    
End Function

Function Transportadora_Exclui(objTransportadora As ClassTransportadora) As Long
'Exclui a Transportadora

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Transportadora_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSPORTADORA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45366
    
    
    Transportadora_Exclui = X.Transportadora_Exclui(objTransportadora)
    
    Exit Function
    
Erro_Transportadora_Exclui:

    Transportadora_Exclui = Err
    
    Select Case Err

        Case 45366
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145459)
        
    End Select
    
    Exit Function
    
End Function

Function TransfCartCobr_Grava1(objTransfCartCobr As ClassTransfCartCobr, iCobradorAnterior As Integer, iCarteiraAnterior As Integer, dValor As Double) As Long
'registra a transferencia de uma parcela de uma carteira de cobranca para outra
'TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TransfCartCobr_Grava1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSFCARTCOBR_GRAVA1", X)
    If lErro <> SUCESSO Then Error 45404
    
    
    TransfCartCobr_Grava1 = X.TransfCartCobr_Grava1(objTransfCartCobr, iCobradorAnterior, iCarteiraAnterior, dValor)
    
    Exit Function
    
Erro_TransfCartCobr_Grava1:

    TransfCartCobr_Grava1 = Err
    
    Select Case Err

        Case 45404
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145460)
        
    End Select
    
    Exit Function
    
End Function


Function PedidosDeVenda_Le_NaturezaOp(sCodigo As String, iFilialEmpresa As Integer, lCodPedVenda As Long) As Long
'funcao pesquisa na tabela de Pedido de Venda a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosDeVenda_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PEDIDOSDEVENDA_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45583
    
    
    PedidosDeVenda_Le_NaturezaOp = X.PedidosDeVenda_Le_NaturezaOp(sCodigo, iFilialEmpresa, lCodPedVenda)
    
    Exit Function
    
Erro_PedidosDeVenda_Le_NaturezaOp:

    PedidosDeVenda_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45583
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145461)
        
    End Select
    
    Exit Function
    
End Function

Function PedidosDeVendaBaixado_Le_NaturezaOp(sCodigo As String, iFilialEmpresa As Integer, lCodPedVenda As Long) As Long
'funcao pesquisa na tabela de Pedido de Venda Baixado a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosDeVendaBaixado_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PEDIDODEVENDABAIXADO_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45584
    
    
    PedidosDeVendaBaixado_Le_NaturezaOp = X.PedidosDeVendaBaixado_Le_NaturezaOp(sCodigo, iFilialEmpresa, lCodPedVenda)
    
    Exit Function
    
Erro_PedidosDeVendaBaixado_Le_NaturezaOp:

    PedidosDeVendaBaixado_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45584
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145462)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_NaturezaOp(sCodigo As String, iFilialEmpresa As Integer, lNumNotaFiscal As Long, sSerie As String) As Long
'funcao pesquisa na tabela de NFiscal a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45585
    
    
    NFiscal_Le_NaturezaOp = X.NFiscal_Le_NaturezaOp(sCodigo, iFilialEmpresa, lNumNotaFiscal, sSerie)
    
    Exit Function
    
Erro_NFiscal_Le_NaturezaOp:

    NFiscal_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45585
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145463)
        
    End Select
    
    Exit Function
    
End Function

' mario 2/7/99 Remover NFiscalBaixada
'Function NFiscalBaixadas_Le_NaturezaOp(sCodigo As String) As Long
''funcao pesquisa na tabela de NFiscal Baixado a natureza passada se encontrou ---> ERRO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscalBaixadas_Le_NaturezaOp
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCALBAIXADAS_LE_NATUREZAOP",X)
'    If lErro <> SUCESSO Then Error 45586
'
'
'    NFiscalBaixadas_Le_NaturezaOp = X.NFiscalBaixadas_Le_NaturezaOp(sCodigo)
'
'    Exit Function
'
'Erro_NFiscalBaixadas_Le_NaturezaOp:
'
'    NFiscalBaixadas_Le_NaturezaOp = Err
'
'    Select Case Err
'
'        Case 45586
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145464)
'
'    End Select
'
'    Exit Function
'
'End Function

Function PadraoTribSaida_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de PadraoTribSaida a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribSaida_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOTRIBSAIDA_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45587
    
    
    PadraoTribSaida_Le_NaturezaOp = X.PadraoTribSaida_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_PadraoTribSaida_Le_NaturezaOp:

    PadraoTribSaida_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45587
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145465)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribEntrada_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de  PadraoTribEntrada a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribEntrada_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PADRAOTRIBENTRADA_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45588
    
    
    PadraoTribEntrada_Le_NaturezaOp = X.PadraoTribEntrada_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_PadraoTribEntrada_Le_NaturezaOp:

    PadraoTribEntrada_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45588
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145466)
        
    End Select
    
    Exit Function
    
End Function

Function TipoDocInfo_Le_NaturezaOp(sCodigo As String, sSigla As String) As Long
'funcao pesquisa na tabela de TipoDocInfo a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPODOCINFO_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45589
    
    
    TipoDocInfo_Le_NaturezaOp = X.TipoDocInfo_Le_NaturezaOp(sCodigo, sSigla)
    
    Exit Function
    
Erro_TipoDocInfo_Le_NaturezaOp:

    TipoDocInfo_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45589
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145467)
        
    End Select
    
    Exit Function
    
End Function

Function TributacaoNF_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoNF a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoNF_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRIBUTACAONF_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45590
    
    
    TributacaoNF_Le_NaturezaOp = X.TributacaoNF_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_TributacaoNF_Le_NaturezaOp:

    TributacaoNF_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45590
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145468)
        
    End Select
    
    Exit Function
    
End Function

Function TributacaoItemPV_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoItemPV a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoItemPV_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRIBUTACAOITEMPV_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45591
    
    
    TributacaoItemPV_Le_NaturezaOp = X.TributacaoItemPV_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_TributacaoItemPV_Le_NaturezaOp:

    TributacaoItemPV_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45591
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145469)
        
    End Select
    
    Exit Function
    
End Function

Function TributacaoComplNF_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoComplNF a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoComplNF_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRIBUTACAOCOMPLNF_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45592
    
    
    TributacaoComplNF_Le_NaturezaOp = X.TributacaoComplNF_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_TributacaoComplNF_Le_NaturezaOp:

    TributacaoComplNF_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45592
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145470)
        
    End Select
    
    Exit Function
    
End Function

Function TributacaoComplPV_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoComplPV a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoComplPV_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRIBUTACAOCOMPLPV_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45593
    
    
    TributacaoComplPV_Le_NaturezaOp = X.TributacaoComplPV_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_TributacaoComplPV_Le_NaturezaOp:

    TributacaoComplPV_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45593
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145471)
        
    End Select
    
    Exit Function
    
End Function

Function TributacaoItemNF_Le_NaturezaOp(sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoItemNF a natureza passada se encontrou ---> ERRO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoItemNF_Le_NaturezaOp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRIBUTACAOITEMNF_LE_NATUREZAOP", X)
    If lErro <> SUCESSO Then Error 45594
    
    
    TributacaoItemNF_Le_NaturezaOp = X.TributacaoItemNF_Le_NaturezaOp(sCodigo)
    
    Exit Function
    
Erro_TributacaoItemNF_Le_NaturezaOp:

    TributacaoItemNF_Le_NaturezaOp = Err
    
    Select Case Err

        Case 45594
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145472)
        
    End Select
    
    Exit Function
    
End Function

Function Serie_Lock1(sSerie As String) As Long
'Faz lock da Série passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Lock1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_SERIE_LOCK1", X)
    If lErro <> SUCESSO Then Error 45595
    
    
    Serie_Lock1 = X.Serie_Lock1(sSerie)
    
    Exit Function
    
Erro_Serie_Lock1:

    Serie_Lock1 = Err
    
    Select Case Err

        Case 45595
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145473)
        
    End Select
    
    Exit Function
    
End Function

Function Transportadora_Lock1(iTransportadora As Integer) As Long
'Faz lock da Transportadora passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Transportadora_Lock1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TRANSPORTADORA_LOCK1", X)
    If lErro <> SUCESSO Then Error 45596
    
    
    Transportadora_Lock1 = X.Transportadora_Lock1(iTransportadora)
    
    Exit Function
    
Erro_Transportadora_Lock1:

    Transportadora_Lock1 = Err
    
    Select Case Err

        Case 45596
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145474)
        
    End Select
    
    Exit Function
    
End Function


Function TabelaPreco_Altera(objTabelaPreco As ClassTabelaPreco) As Long
'Atualiza Tabela de Preço existente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELAPRECO_ALTERA", X)
    If lErro <> SUCESSO Then Error 45598
    
    
    TabelaPreco_Altera = X.TabelaPreco_Altera(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPreco_Altera:

    TabelaPreco_Altera = Err
    
    Select Case Err

        Case 45598
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145475)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPreco_Cria(objTabelaPreco As ClassTabelaPreco) As Long
'Cria uma nova Tabela de Preço

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Cria
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELAPRECO_CRIA", X)
    If lErro <> SUCESSO Then Error 45599
    
    
    TabelaPreco_Cria = X.TabelaPreco_Cria(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPreco_Cria:

    TabelaPreco_Cria = Err
    
    Select Case Err

        Case 45599
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145476)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPreco_Automatico(lNumProxCodigo As Long) As Long
'Gera o próximo código da Tabela de Preço automático.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELAPRECO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45600
    
    
    TabelaPreco_Automatico = X.TabelaPreco_Automatico(lNumProxCodigo)
    
    Exit Function
    
Erro_TabelaPreco_Automatico:

    TabelaPreco_Automatico = Err
    
    Select Case Err

        Case 45600
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145477)
        
    End Select
    
    Exit Function
    
End Function

Function CRFATConfig_ObterNumInt1(sCodigo As String, lNumInt As Long, Optional iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_ObterNumInt1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CRFATCONFIG_OBTERNUMINT1", X)
    If lErro <> SUCESSO Then Error 45682
    
    
    CRFATConfig_ObterNumInt1 = X.CRFATConfig_ObterNumInt1(sCodigo, lNumInt, iQtde)
    
    Exit Function
    
Erro_CRFATConfig_ObterNumInt1:

    CRFATConfig_ObterNumInt1 = Err
    
    Select Case Err

        Case 45682
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145478)
        
    End Select
    
    Exit Function
    
End Function

'mario 2/7/99 Remover NFiscalBaixada
'Function ItemNFiscalBaixada_Le(objItemNFiscalBaixada As ClassItemNF) As Long
''Lê em ItensNFiscalBaixada o Ítem com o Número Interno passado.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ItemNFiscalBaixada_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_ITEMNFISCALBAIXADA_LE",X)
'    If lErro <> SUCESSO Then Error 45605
'
'
'    ItemNFiscalBaixada_Le = X.ItemNFiscalBaixada_Le(objItemNFiscalBaixada)
'
'    Exit Function
'
'Erro_ItemNFiscalBaixada_Le:
'
'    ItemNFiscalBaixada_Le = Err
'
'    Select Case Err
'
'        Case 45605
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145479)
'
'    End Select
'
'    Exit Function
'
'End Function

Function ItemNFiscal_Le(objItemNFiscal As ClassItemNF) As Long
'Lê o Ítem de Nota Fiscal com o Número Interno passado em objItemNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFiscal_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITEMNFISCAL_LE", X)
    If lErro <> SUCESSO Then Error 45604
    
    
    ItemNFiscal_Le = X.ItemNFiscal_Le(objItemNFiscal)
    
    Exit Function
    
Erro_ItemNFiscal_Le:

    ItemNFiscal_Le = Err
    
    Select Case Err

        Case 45604
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145480)
        
    End Select
    
    Exit Function
    
End Function

'mario
'Function NFiscalBaixada_Le(objNFiscal As ClassNFiscal) As Long
''Lê a Nota Fiscal com o Número Interno passado em objNFiscal.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscalBaixada_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALBAIXADA_LE",X)
'    If lErro <> SUCESSO Then Error 44375
'
'
'    NFiscalBaixada_Le = X.NFiscalBaixada_Le(objNFiscal)
'
'    Exit Function
'
'Erro_NFiscalBaixada_Le:
'
'    NFiscalBaixada_Le = Err
'
'    Select Case Err
'
'        Case 44375
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145481)
'
'    End Select
'
'    Exit Function
'
'End Function

' mario 2/7/99 Remover NFiscalBaixada
'Function NFiscalBaixadaItens_Le(objNFiscal As ClassNFiscal) As Long
''Lê em ItensNFiscalBaixadas os ítens da Nota Fiscal passada em objNFiscal
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscalBaixadaItens_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALBAIXADAITENS_LE",X)
'    If lErro <> SUCESSO Then Error 45603
'
'
'    NFiscalBaixadaItens_Le = X.NFiscalBaixadaItens_Le(objNFiscal)
'
'    Exit Function
'
'Erro_NFiscalBaixadaItens_Le:
'
'    NFiscalBaixadaItens_Le = Err
'
'    Select Case Err
'
'        Case 45603
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145482)
'
'    End Select
'
'    Exit Function
'
'End Function

Function NFiscalItens_Le(objNFiscal As ClassNFiscal) As Long
'Lê os ítens da Nota Fiscal passada em objNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALITENS_LE", X)
    If lErro <> SUCESSO Then Error 19697
    
    
    NFiscalItens_Le = X.NFiscalItens_Le(objNFiscal)

    Exit Function
    
Erro_NFiscalItens_Le:
    
    NFiscalItens_Le = Err
    
    Select Case Err

        Case 19697
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145483)
        
    End Select
    
    Exit Function

End Function

Function NFiscalInternaSaida_Le_Numero(objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'          NomeTelaNFiscal<>"NFiscalDev"
'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalInternaSaida_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALINTERNASAIDA_LE_NUMERO", X)
    If lErro <> SUCESSO Then Error 45582
    
    
    NFiscalInternaSaida_Le_Numero = X.NFiscalInternaSaida_Le_Numero(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalInternaSaida_Le_Numero:

    NFiscalInternaSaida_Le_Numero = Err
    
    Select Case Err

        Case 45582
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145484)
        
    End Select
    
    Exit Function
    
End Function

Function TabelasPreco_Le_Codigos(colCodigo As Collection) As Long
'Lê todos os códigos contidos na Tabela De Preços e coloca-os em colCodigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelasPreco_Le_Codigos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELASPRECO_LE_CODIGOS", X)
    If lErro <> SUCESSO Then Error 45745
    
    
    TabelasPreco_Le_Codigos = X.TabelasPreco_Le_Codigos(colCodigo)
    
    Exit Function
    
Erro_TabelasPreco_Le_Codigos:

    TabelasPreco_Le_Codigos = Err
    
    Select Case Err

        Case 45745
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145485)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPreco_Le(objTabelaPreco As ClassTabelaPreco) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TABELAPRECO_LE", X)
    If lErro <> SUCESSO Then Error 45746
    
    
    TabelaPreco_Le = X.TabelaPreco_Le(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPreco_Le:

    TabelaPreco_Le = Err
    
    Select Case Err

        Case 45746
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145486)
        
    End Select
    
    Exit Function
    
End Function


Function FilialEmpresa_Le_Codigos(colCodigo As Collection) As Long
'Lê todos os códigos de Filiais da Empresa corrente na tabela FiliaisEmpresa.
'Devolve em colCodFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialEmpresa_Le_Codigos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALEMPRESA_LE_CODIGOS", X)
    If lErro <> SUCESSO Then Error 45748
    
    
    FilialEmpresa_Le_Codigos = X.FilialEmpresa_Le_Codigos(colCodigo)
    
    Exit Function
    
Erro_FilialEmpresa_Le_Codigos:

    FilialEmpresa_Le_Codigos = Err
    
    Select Case Err

        Case 45748
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145487)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPreco_Exclui(objTabelaPreco As ClassTabelaPreco) As Long
'Exclui a Tabela e Preço do BD
'??? criar uma funcao em separado para testar se pode excluir

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELAPRECO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45749
    
    
    TabelaPreco_Exclui = X.TabelaPreco_Exclui(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPreco_Exclui:

    TabelaPreco_Exclui = Err
    
    Select Case Err

        Case 45749
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145488)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItens_Exclui(objTabelaPreco As ClassTabelaPreco) As Long
'Rotina responsável pela exclusão de itens da Tabela de Preço relacionados à TabelasDePreco
'DEVE SE CHAMADA DENTRO DE UMA TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItens_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELAPRECOITENS_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45750
    
    
    TabelaPrecoItens_Exclui = X.TabelaPrecoItens_Exclui(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPrecoItens_Exclui:

    TabelaPrecoItens_Exclui = Err
    
    Select Case Err

        Case 45750
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145489)
        
    End Select
    
    Exit Function
    
End Function



Function TabelaPreco_Exclui_Precos(objTabelaPreco As ClassTabelaPreco) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPreco_Exclui_Precos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELAPRECO_EXCLUI_PRECOS", X)
    If lErro <> SUCESSO Then Error 45753
    
    
    TabelaPreco_Exclui_Precos = X.TabelaPreco_Exclui_Precos(objTabelaPreco)
    
    Exit Function
    
Erro_TabelaPreco_Exclui_Precos:

    TabelaPreco_Exclui_Precos = Err
    
    Select Case Err

        Case 45753
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145490)
        
    End Select
    
    Exit Function
    
End Function

Function Clientes_Le_TabelaPreco(objTabelaPreco As ClassTabelaPreco, colCodCliente As Collection) As Long
'Lê códigos de Clientes que se relacionam com objTabelaPreco e coloca-os em colCodClientes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Clientes_Le_TabelaPreco
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CLIENTES_LE_TABELAPRECO", X)
    If lErro <> SUCESSO Then Error 45754
    
    
    Clientes_Le_TabelaPreco = X.Clientes_Le_TabelaPreco(objTabelaPreco, colCodCliente)
    
    Exit Function
    
Erro_Clientes_Le_TabelaPreco:

    Clientes_Le_TabelaPreco = Err
    
    Select Case Err

        Case 45754
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145491)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasRec_Inserir(objParcelaReceber As ClassParcelaReceber, lComandoParcIns As Long) As Long
'auxiliar ao cancelamento de baixa de parcela a receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasRec_Inserir
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASREC_INSERIR", X)
    If lErro <> SUCESSO Then Error 45835
    
    
    ParcelasRec_Inserir = X.ParcelasRec_Inserir(objParcelaReceber, lComandoParcIns)
    
    Exit Function
    
Erro_ParcelasRec_Inserir:

    ParcelasRec_Inserir = Err
    
    Select Case Err

        Case 45835
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145492)
        
    End Select
    
    Exit Function
    
End Function

Function CRFATConfig_Le_Configs(colConfigs As ColCRFATConfig) As Long
'recupera da tabela FATConfig uma colecao de registros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_Le_Configs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CRFATCONFIG_LE_CONFIGS", X)
    If lErro <> SUCESSO Then Error 55006
    
    
    CRFATConfig_Le_Configs = X.CRFATConfig_Le_Configs(colConfigs)
    
    Exit Function
    
Erro_CRFATConfig_Le_Configs:

    CRFATConfig_Le_Configs = Err
    
    Select Case Err

        Case 55006
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145493)
        
    End Select
    
    Exit Function
    
End Function

Function CRFATConfig_Grava_Configs_Trans(colConfigs As ColCRFATConfig) As Long
'grava na tabela CRFATConfig uma colecao de registros de configuracao
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_Grava_Configs_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CRFATCONFIG_GRAVA_CONFIGS_TRANS", X)
    If lErro <> SUCESSO Then Error 55018
    
    
    CRFATConfig_Grava_Configs_Trans = X.CRFATConfig_Grava_Configs_Trans(colConfigs)
    
    Exit Function
    
Erro_CRFATConfig_Grava_Configs_Trans:

    CRFATConfig_Grava_Configs_Trans = Err
    
    Select Case Err

        Case 55018
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145494)
        
    End Select
    
    Exit Function
    
End Function

'??? fernando, criar chama funcoes. Esta funcao esta em CRFATGrava.cls
Function CRFATConfig_Grava_Configs(colConfigs As ColCRFATConfig) As Long
'grava na tabela CRFATConfig uma colecao de registros de configuracao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_Grava_Configs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CRFATCONFIG_GRAVA_CONFIGS", X)
    If lErro <> SUCESSO Then Error 54540
    
    
    CRFATConfig_Grava_Configs = X.CRFATConfig_Grava_Configs(colConfigs)
    
    Exit Function
    
Erro_CRFATConfig_Grava_Configs:

    CRFATConfig_Grava_Configs = Err
    
    Select Case Err

        Case 54540
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145495)
        
    End Select
    
    Exit Function
    
End Function
Function ChequePre_Grava(objChequePre As ClassChequePre, colInfoParcRec As Collection) As Long
'Rotina que grava o Cheque Pre

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CHEQUEPRE_GRAVA", X)
    If lErro <> SUCESSO Then Error 27014
    
    
    ChequePre_Grava = X.ChequePre_Grava(objChequePre, colInfoParcRec)
    
    Exit Function
    
Erro_ChequePre_Grava:

    ChequePre_Grava = Err
    
    Select Case Err

        Case 27014
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145496)
        
    End Select
    
    Exit Function
    
End Function

Function ChequePre_Exclui(objChequePre As ClassChequePre) As Long
'Rotina responsavel pela exclusao de Cheque Pre

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CHEQUEPRE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 27015
    
    
    ChequePre_Exclui = X.ChequePre_Exclui(objChequePre)
    
    Exit Function
    
Erro_ChequePre_Exclui:

    ChequePre_Exclui = Err
    
    Select Case Err

        Case 27015
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145497)
        
    End Select
    
    Exit Function
    
End Function

Function DocumentoCPR_OrigemNFiscal(sSigla As String, lNumIntDoc As Long, lNumIntNF As Long) As Long
'retorna SUCESSO se existe registro em NFiscal associado a um docto em CPR identificado pela sigla conforme a tabela TiposDeDocumento e lNumIntDoc na tabela em CPR
'e retorna 41542 se nao existir
'se existir isto quer dizer que o documento em CPR foi criado atraves do registro da nota fiscal nos modulos de faturamento ou estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DocumentoCPR_OrigemNFiscal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_DOCUMENTOCPR_ORIGEMNFISCAL", X)
    If lErro <> SUCESSO Then Error 54718
    
    
    DocumentoCPR_OrigemNFiscal = X.DocumentoCPR_OrigemNFiscal(sSigla, lNumIntDoc, lNumIntNF)
    
    Exit Function
    
Erro_DocumentoCPR_OrigemNFiscal:

    DocumentoCPR_OrigemNFiscal = Err
    
    Select Case Err

        Case 54718
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145498)
        
    End Select
    
    Exit Function
    
End Function

Function Parcelas_Calcula(dValor As Double, iNumParcelas As Integer, colValorParcelas As Collection) As Long
'Calcula os valores das Parcelas a partir do Valor e Número de Parcelas
'Devolve em colValorParcelas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Parcelas_Calcula
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELAS_CALCULA", X)
    If lErro <> SUCESSO Then Error 54720
    
    
    Parcelas_Calcula = X.Parcelas_Calcula(dValor, iNumParcelas, colValorParcelas)
    
    Exit Function
    
Erro_Parcelas_Calcula:

    Parcelas_Calcula = Err
    
    Select Case Err

        Case 54720
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145499)
        
    End Select
    
    Exit Function
    
End Function

Function Parcelas_DatasVencimento(objCondicaoPagto As ClassCondicaoPagto, dtDataEmissao As Date, colDataVencimento As Collection) As Long
'Calcula as datas de vencimento de cada uma das Parcelas do Grid

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Parcelas_DatasVencimento
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELAS_DATASVENCIMENTO", X)
    If lErro <> SUCESSO Then Error 54721
    
    
    Parcelas_DatasVencimento = X.Parcelas_DatasVencimento(objCondicaoPagto, dtDataEmissao, colDataVencimento)
    
    Exit Function
    
Erro_Parcelas_DatasVencimento:

    Parcelas_DatasVencimento = Err
    
    Select Case Err

        Case 54721
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145500)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceberBaixado_Le(objTituloReceber As ClassTituloReceber, Optional iExcluido As Integer = 0) As Long
'Lê Titulo a Receber Baixado com o Número Interno passado como parâmetro
'No Parametro Optional: Se iExcluido = 1 então entra os Titulos Excluidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceberBaixado_Le
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBERBAIXADO_LE", X)
    If lErro <> SUCESSO Then Error 54723
    
    TituloReceberBaixado_Le = X.TituloReceberBaixado_Le(objTituloReceber, iExcluido)
    
    Exit Function
    
Erro_TituloReceberBaixado_Le:

    TituloReceberBaixado_Le = Err
    
    Select Case Err

        Case 54723
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145501)
        
    End Select
    
    Exit Function
    
End Function

Function Parcela_GeraDescontoPadrao(ColDesconto As ColDesconto, dtVencimentoParc As Date) As Long
'preenche a colecao de descontos padrao para a parcela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Parcela_GeraDescontoPadrao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELA_GERADESCONTOPADRAO", X)
    If lErro <> SUCESSO Then Error 54724
    
    
    Parcela_GeraDescontoPadrao = X.Parcela_GeraDescontoPadrao(ColDesconto, dtVencimentoParc)
    
    Exit Function
    
Erro_Parcela_GeraDescontoPadrao:

    Parcela_GeraDescontoPadrao = Err
    
    Select Case Err

        Case 54724
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145502)
        
    End Select
    
    Exit Function
    
End Function

Function Calcula_Desconto_Parcela(objParcelaReceber As ClassParcelaReceber, dValorDesconto As Double, dtDataBaixa As Date) As Long
'Obtem o valor do desconto por antecipacao de pagamento de uma parcela na data informada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Calcula_Desconto_Parcela
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CALCULA_DESCONTO_PARCELA", X)
    If lErro <> SUCESSO Then Error 54752
    
    
    Calcula_Desconto_Parcela = X.Calcula_Desconto_Parcela(objParcelaReceber, dValorDesconto, dtDataBaixa)
    
    Exit Function
    
Erro_Calcula_Desconto_Parcela:

    Calcula_Desconto_Parcela = Err
    
    Select Case Err

        Case 54752
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145503)
        
    End Select
    
    Exit Function
    
End Function

Function Calcula_Multa_Juros_Parcela(objParcelaReceber As ClassParcelaReceber, dtDataBaixa As Date, dValorMulta As Double, dValorJuros As Double) As Long
'Obtem o valor da multa e dos juros para pagto da parcela na data informada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Calcula_Multa_Juros_Parcela
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CALCULA_MULTA_JUROS_PARCELA", X)
    If lErro <> SUCESSO Then Error 54753
    
    
    Calcula_Multa_Juros_Parcela = X.Calcula_Multa_Juros_Parcela(objParcelaReceber, dtDataBaixa, dValorMulta, dValorJuros)
    
    Exit Function
    
Erro_Calcula_Multa_Juros_Parcela:

    Calcula_Multa_Juros_Parcela = Err
    
    Select Case Err

        Case 54753
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145504)
        
    End Select
    
    Exit Function
    
End Function

Function TitulosPag_Le_CondicaoPagto(iCodigo As Integer) As Long
'Testa a Integridade da Titulos a Pagar e Titulos a Pagar Baixados
'com relação a Condição de Pagamento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TitulosPag_Le_CondicaoPagto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULOSPAG_LE_CONDICAOPAGTO", X)
    If lErro <> SUCESSO Then Error 54779
    
    
    TitulosPag_Le_CondicaoPagto = X.TitulosPag_Le_CondicaoPagto(iCodigo)
    
    Exit Function
    
Erro_TitulosPag_Le_CondicaoPagto:

    TitulosPag_Le_CondicaoPagto = Err
    
    Select Case Err

        Case 54779
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145505)
        
    End Select
    
    Exit Function
    
End Function

Function TipoCliente_Transportadora_Nao_Relacionada(iCodigo As Integer) As Long
'Verifica se a Transportadora identificada por iCodigo está relacionada a algum Tipo de Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoCliente_Transportadora_Nao_Relacionada
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TIPOCLIENTE_TRANSP_NAO_RELACIONADA", X)
    If lErro <> SUCESSO Then Error 54780
    
    
    TipoCliente_Transportadora_Nao_Relacionada = X.TipoCliente_Transportadora_Nao_Relacionada(iCodigo)
    
    Exit Function
    
Erro_TipoCliente_Transportadora_Nao_Relacionada:

    TipoCliente_Transportadora_Nao_Relacionada = Err
    
    Select Case Err

        Case 54780
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145506)
        
    End Select
    
    Exit Function
    
End Function

Function TitulosRec_Le_CondicaoPagto(iCodigo As Integer) As Long
'Testa a Integridade da Titulos a Receber e Titulos a Receber Baixados
'com relação a Condição de Pagamento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TitulosRec_Le_CondicaoPagto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULOSREC_LE_CONDICAOPAGTO", X)
    If lErro <> SUCESSO Then Error 54781
    
    
    TitulosRec_Le_CondicaoPagto = X.TitulosRec_Le_CondicaoPagto(iCodigo)
    
    Exit Function
    
Erro_TitulosRec_Le_CondicaoPagto:

    TitulosRec_Le_CondicaoPagto = Err
    
    Select Case Err

        Case 54781
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145507)
        
    End Select
    
    Exit Function
    
End Function

Function ChequePre_Le_FilialCliente(objFilialCliente As ClassFilialCliente) As Long
'Lê na tabela de ChequePre, FilialCliente à partir da chave(lCodigo) fornecida em objFilialCliente.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ChequePre_Le_FilialCliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CHEQUEPRE_LE_FILIALCLIENTE", X)
    If lErro <> SUCESSO Then Error 54782
    
    
    ChequePre_Le_FilialCliente = X.ChequePre_Le_FilialCliente(objFilialCliente)
    
    Exit Function
    
Erro_ChequePre_Le_FilialCliente:

    ChequePre_Le_FilialCliente = Err
    
    Select Case Err

        Case 54782
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145508)
        
    End Select
    
    Exit Function
    
End Function

Public Function ParcelasPag_Le_DataValor(lNumIntTitulo As Long, colParcPag As colParcelaPagar, lComando As Long) As Long
'Lê as datas de vencimento e os valores das ParcelasPag associadas ao título com NumIntTitulo passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPag_Le_DataValor
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASPAG_LE_DATAVALOR", X)
    If lErro <> SUCESSO Then Error 54812
    
    
    ParcelasPag_Le_DataValor = X.ParcelasPag_Le_DataValor(lNumIntTitulo, colParcPag, lComando)
    
    Exit Function
    
Erro_ParcelasPag_Le_DataValor:

    ParcelasPag_Le_DataValor = Err
    
    Select Case Err

        Case 54812
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145509)
        
    End Select
    
    Exit Function
    
End Function

Public Function ParcelasRec_Le_DataValor(lNumIntTitulo As Long, colParcRec As colParcelaReceber, lComando As Long) As Long
'Lê as datas de vencimento e os valores das ParcelasRec associadas ao título com NumIntTitulo passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasRec_Le_DataValor
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELASREC_LE_DATAVALOR", X)
    If lErro <> SUCESSO Then Error 54813
    
    
    ParcelasRec_Le_DataValor = X.ParcelasRec_Le_DataValor(lNumIntTitulo, colParcRec, lComando)
    
    Exit Function
    
Erro_ParcelasRec_Le_DataValor:

    ParcelasRec_Le_DataValor = Err
    
    Select Case Err

        Case 54813
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145510)
        
    End Select
    
    Exit Function
    
End Function

Function ClienteCategorias_Le_FilialCliente(objFilialCliente As ClassFilialCliente, colCateg As Collection) As Long
'preenche uma colecao com objClienteCategoria(s) associados a objFilialCliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClienteCategorias_Le_FilialCliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CLIENTECATEGORIAS_LE_FILIALCLIENTE", X)
    If lErro <> SUCESSO Then Error 54814
    
    
    ClienteCategorias_Le_FilialCliente = X.ClienteCategorias_Le_FilialCliente(objFilialCliente, colCateg)
    
    Exit Function
    
Erro_ClienteCategorias_Le_FilialCliente:

    ClienteCategorias_Le_FilialCliente = Err
    
    Select Case Err

        Case 54814
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145511)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesNF_Gera_ComissoesParc(objNFiscal As ClassNFiscal, colcolComissao As colcolComissao, dValorLiqTitulo As Double) As Long
'Preenche a colecao de colecoes de comissoes a nivel de parcela de um titulo a receber associado a NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesNF_Gera_ComissoesParc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISSOESNF_GERA_COMISSOESPARC", X)
    If lErro <> SUCESSO Then Error 54872
    
    
    ComissoesNF_Gera_ComissoesParc = X.ComissoesNF_Gera_ComissoesParc(objNFiscal, colcolComissao, dValorLiqTitulo)
    
    Exit Function
    
Erro_ComissoesNF_Gera_ComissoesParc:

    ComissoesNF_Gera_ComissoesParc = Err
    
    Select Case Err

        Case 54872
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145512)
        
    End Select
    
    Exit Function
    
End Function


Function ParcelaReceber_Le(objParcelaReceber As ClassParcelaReceber) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_ParcelaReceber_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_PARCELARECEBER_LE", X)
    If lErro <> SUCESSO Then Error 45389
    
    
    ParcelaReceber_Le = X.ParcelaReceber_Le(objParcelaReceber)
    
    Exit Function
    
Erro_ParcelaReceber_Le:

    ParcelaReceber_Le = Err
    
    Select Case Err

        Case 45389
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145513)
        
    End Select
    
    Exit Function
    
End Function

Public Function ComissaoAutomatica_Obter_Parcela(objVendedor As ClassVendedor, dValorParcela As Double, objComissaoNF As ClassComissaoNF, objComissao As ClassComissao) As Long
'retorna em objComissao valores default para geracao de comissao automatica à partir do cliente, filial e valor da parcela informados
'se objVendedor.iCodigo <> 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim X As Object, lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_Parcela
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_AUTO_OBTER_PARC", X)
    If lErro <> SUCESSO Then Error 59030
    
    
    ComissaoAutomatica_Obter_Parcela = X.ComissaoAutomatica_Obter_Parcela(objVendedor, dValorParcela, objComissaoNF, objComissao)
    
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Parcela:

    ComissaoAutomatica_Obter_Parcela = Err
    
    Select Case Err

        Case 59030
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145514)
        
    End Select
    
    Exit Function
    
End Function

Public Function ComissaoAutomatica_Obter_Emissao(sClienteNomeRed As String, iCodFilial As Integer, objVendedor As ClassVendedor, objComissao As ClassComissao, dValorTotal As Double, dValorFrete As Double, dValorSeguro As Double, dValorOutras As Double, dValorDesconto As Double) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim X As Object, lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_Emissao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_AUTO_OBTER_EMISS", X)
    If lErro <> SUCESSO Then Error 59031
    
    
    ComissaoAutomatica_Obter_Emissao = X.ComissaoAutomatica_Obter_Emissao(sClienteNomeRed, iCodFilial, objVendedor, objComissao, dValorTotal, dValorFrete, dValorSeguro, dValorOutras, dValorDesconto)
    
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Emissao:

    ComissaoAutomatica_Obter_Emissao = Err
    
    Select Case Err

        Case 59031
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145515)
        
    End Select
    
    Exit Function
    
End Function

Public Function ComissaoAutomatica_Obter_Info(sClienteNomeRed As String, iCodFilial As Integer, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim X As Object, lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_Info
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_AUTO_OBTER_INFO", X)
    If lErro <> SUCESSO Then Error 59032
    
    
    ComissaoAutomatica_Obter_Info = X.ComissaoAutomatica_Obter_Info(sClienteNomeRed, iCodFilial, objVendedor, objComissaoNF)
    
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Info:

    ComissaoAutomatica_Obter_Info = Err
    
    Select Case Err

        Case 59032
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145516)
        
    End Select
    
    Exit Function
    
End Function

Public Function ComissaoAutomatica_Obter_Info2(objFilialCliente As ClassFilialCliente, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim X As Object, lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_Info2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_AUTO_OBTER_INFO2", X)
    If lErro <> SUCESSO Then Error 59033
    
    
    ComissaoAutomatica_Obter_Info2 = X.ComissaoAutomatica_Obter_Info2(objFilialCliente, objVendedor, objComissaoNF)
    
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Info2:

    ComissaoAutomatica_Obter_Info2 = Err
    
    Select Case Err

        Case 59033
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145517)
        
    End Select
    
    Exit Function
    
End Function

Public Function ComissaoAutomatica_Obter_NF(objFilialCliente As ClassFilialCliente, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF, dValorTotal As Double, dValorFrete As Double, dValorSeguro As Double, dValorOutras As Double, dValorDesconto As Double) As Long
'retorna em objComissao valores default para geracao de comissao automatica
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim X As Object, lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_NF
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_AUTO_OBTER_NF", X)
    If lErro <> SUCESSO Then Error 59034
    
    
    ComissaoAutomatica_Obter_NF = X.ComissaoAutomatica_Obter_NF(objFilialCliente, objVendedor, objComissaoNF, dValorTotal, dValorFrete, dValorSeguro, dValorOutras, dValorDesconto)
    
    Exit Function
    
Erro_ComissaoAutomatica_Obter_NF:

    ComissaoAutomatica_Obter_NF = Err
    
    Select Case Err

        Case 59034
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145518)
        
    End Select
    
    Exit Function
    
End Function

Public Function Comissao_GridFAT_Valida(objGridComissoes As AdmGrid, iGrid_Vendedor_Col As Integer, iGrid_ValorCom_Col As Integer, dValorTotal As Double, dValorIRRF As Double) As Long
'verifica se o grid de comissoes está preenchido com valores válidos

Dim X As Object, lErro As Long

On Error GoTo Erro_Comissao_GridFAT_Valida
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_COMISS_VALIDA_GRID", X)
    If lErro <> SUCESSO Then Error 59035
    
    
    Comissao_GridFAT_Valida = X.Comissao_GridFAT_Valida(objGridComissoes, iGrid_Vendedor_Col, iGrid_ValorCom_Col, dValorTotal, dValorIRRF)
    
    Exit Function
    
Erro_Comissao_GridFAT_Valida:

    Comissao_GridFAT_Valida = Err
    
    Select Case Err

        Case 59035
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145519)
        
    End Select
    
    Exit Function
    
End Function


Function DebitoReceber_Le(objDebitoRecCli As ClassDebitoRecCli, Optional iExclusao As Integer = 0) As Long
'Le o debito a receber a partir da chave fornecida em objDebitoRecCli. Retorna os dados em objDebitosRecCli

Dim X As Object, lErro As Long
    
On Error GoTo Erro_DebitoReceber_Le
    
    lErro = Rotina_ObterObjetoFuncao("CPR_DEBITORECEBER_LE", X)
    If lErro <> SUCESSO Then Error 19969
    
    DebitoReceber_Le = X.DebitoReceber_Le(objDebitoRecCli, iExclusao)
    
    Exit Function
    
Erro_DebitoReceber_Le:

    DebitoReceber_Le = Err
    
    Select Case Err

        Case 19969
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145520)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelaReceber_Baixada_Le(objParcelaReceber As ClassParcelaReceber) As Long
'le dados de uma parcela Baixada de um titulo a receber identificada por objParcelaReceber.NumIntDoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelaReceber_Baixada_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PARCELARECEBER_BAIXADA_LE", X)
    If lErro <> SUCESSO Then Error 54922
    
    
    ParcelaReceber_Baixada_Le = X.ParcelaReceber_Baixada_Le(objParcelaReceber)
    
    Exit Function
    
Erro_ParcelaReceber_Baixada_Le:

    ParcelaReceber_Baixada_Le = Err
    
    Select Case Err

        Case 54922
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145521)
        
    End Select
    
    Exit Function
    
End Function

Function FilialCliente_Testa_CGC(lCodCliente As Long, iCodFilial As Integer, sCgc As String) As Long
'Verifica se tem outro cliente com o mesmo CGC e dá aviso.
'Não está levando em conta Filial, será melhorado futuramente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Testa_CGC
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALCLIENTE_TESTA_CGC", X)
    If lErro <> SUCESSO Then Error 54997
    
    
    FilialCliente_Testa_CGC = X.FilialCliente_Testa_CGC(lCodCliente, iCodFilial, sCgc)
    
    Exit Function
    
Erro_FilialCliente_Testa_CGC:

    FilialCliente_Testa_CGC = Err
    
    Select Case Err

        Case 54997
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145522)
        
    End Select
    
    Exit Function
    
End Function

Function Cobrador_Le(objCobrador As ClassCobrador) As Long
'le o Cobrador a partir da chave fornecida em objCobrador. Devolve os dados em objCobrador

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cobrador_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_COBRADOR_LE", X)
    If lErro <> SUCESSO Then Error 14585
    
    
    Cobrador_Le = X.Cobrador_Le(objCobrador)
    
    Exit Function
    
Erro_Cobrador_Le:

    Cobrador_Le = Err
    
    Select Case Err

        Case 14585
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145523)
        
    End Select
    
    Exit Function
    
End Function

Function Cobrador_Le_NomeReduzido(objCobrador As ClassCobrador) As Long
'Lê os dados do Cobrador à partir do Nome Reduzido passado como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_Cobrador_Le_NomeReduzido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_COBRADOR_LE_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then Error 45349
    
    
    Cobrador_Le_NomeReduzido = X.Cobrador_Le_NomeReduzido(objCobrador)
    
    Exit Function
    
Erro_Cobrador_Le_NomeReduzido:

    Cobrador_Le_NomeReduzido = Err
    
    Select Case Err

        Case 45349
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145524)
        
    End Select
    
    Exit Function
    
End Function

Public Function Cliente_Le_Codigo_NomeReduzido(objCliente As ClassCliente) As Long
'obtem o codigo do cliente à partir de seu nome reduzido

Dim X As Object, lErro As Long

On Error GoTo Erro_Cliente_Le_Codigo_NomeReduzido
      
    lErro = Rotina_ObterObjetoFuncao("CPR_CLIENTE_LE_CODIGO_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then Error 59738
      
    Cliente_Le_Codigo_NomeReduzido = X.Cliente_Le_Codigo_NomeReduzido(objCliente)
    
    Exit Function
    
Erro_Cliente_Le_Codigo_NomeReduzido:

    Cliente_Le_Codigo_NomeReduzido = Err
    
    Select Case Err

        Case 59738
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145525)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalInternaSaida_Le_Numero2(objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'          NomeTelaNFiscal<>"NFiscalDev"
'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalInternaSaida_Le_Numero2
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALINTERNASAIDA_LE_NUMERO2", X)
    If lErro <> SUCESSO Then Error 45582
    
    NFiscalInternaSaida_Le_Numero2 = X.NFiscalInternaSaida_Le_Numero2(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalInternaSaida_Le_Numero2:

    NFiscalInternaSaida_Le_Numero2 = Err
    
    Select Case Err

        Case 45582
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145526)
        
    End Select
    
    Exit Function
    
End Function

Function CreditoPagar_Exclui_EmTransacao(objCreditoPagar As ClassCreditoPagar, objContabil As ClassContabil) As Long
'Exclui Crédito a Pagar do BD
'ATENÇÃO: ESSA ROTINA DEVE SER CHAMADA DENTRO DE UMA TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CreditoPagar_Exclui_EmTransacao
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CREDITOPAGAR_EXCLUI_EMTRANSACAO", X)
    If lErro <> SUCESSO Then Error 59871
    
    CreditoPagar_Exclui_EmTransacao = X.CreditoPagar_Exclui_EmTransacao(objCreditoPagar, objContabil)
    
    Exit Function
    
Erro_CreditoPagar_Exclui_EmTransacao:

    CreditoPagar_Exclui_EmTransacao = Err
    
    Select Case Err

        Case 59871
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145527)
        
    End Select
    
    Exit Function
    
End Function

Function TituloReceber_Exclui_EmTransacao(objTituloReceber As ClassTituloReceber, objContabil As ClassContabil) As Long
'Exclui o Título à Receber do Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloReceber_Exclui_EmTransacao
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_TITULORECEBER_EXCLUI_EMTRANSACAO", X)
    If lErro <> SUCESSO Then Error 62140
    
    
    TituloReceber_Exclui_EmTransacao = X.TituloReceber_Exclui_EmTransacao(objTituloReceber, objContabil)
    
    Exit Function
    
Erro_TituloReceber_Exclui_EmTransacao:

    TituloReceber_Exclui_EmTransacao = Err
    
    Select Case Err

        Case 62140
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145528)
        
    End Select
    
    Exit Function
    
End Function

Function CarteiraCobrador_LeMod(objCarteiraCobrador As ClassCarteiraCobrador, lComando As Long) As Long
'faz a leitura de um registro da tabela CarteirasCobrador p/objCarteiraCobrador c/ExecutarPos
'objCarteiraCobrador.iCobrador e objCarteiraCobrador.iCodCarteiraCobranca entram preenchidos
'TEM QUE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CarteiraCobrador_LeMod
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_CARTEIRACOBRADOR_LEMOD", X)
    If lErro <> SUCESSO Then Error 24703
    
    
    CarteiraCobrador_LeMod = X.CarteiraCobrador_LeMod(objCarteiraCobrador, lComando)
    
    Exit Function
    
Erro_CarteiraCobrador_LeMod:

    CarteiraCobrador_LeMod = Err
    
    Select Case Err

        Case 24703
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145529)
        
    End Select
    
    Exit Function
    
End Function


Function CRFatConfig_Le(sCodigo As String, iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de CRFatConfig de acordo com o código passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFatConfig_Le
      
    lErro = Rotina_ObterObjetoFuncao("CPR_CRFATCONFIG_LE", X)
    If lErro <> SUCESSO Then Error 65081
    
    CRFatConfig_Le = X.CRFatConfig_Le(sCodigo, iFilialEmpresa, sConteudo)
    
    Exit Function
    
Erro_CRFatConfig_Le:

    CRFatConfig_Le = Err
    
    Select Case Err

        Case 65081
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145530)
        
    End Select
    
    Exit Function
    
End Function

Function CRFATConfig_Grava(sCodigo As String, iFilialEmpresa As Integer, sConteudo As String) As Long
'grava na tabela CRFATConfig o Conteudo passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CRFATConfig_Grava
      
    lErro = Rotina_ObterObjetoFuncao("CPR_CRFATCONFIG_GRAVA", X)
    If lErro <> SUCESSO Then Error 65082
    
    CRFATConfig_Grava = X.CRFATConfig_Grava(sCodigo, iFilialEmpresa, sConteudo)
    
    Exit Function
    
Erro_CRFATConfig_Grava:

    CRFATConfig_Grava = Err
    
    Select Case Err

        Case 65082
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145531)
        
    End Select
    
    Exit Function
    
End Function

Function NFPag_Exclui_EmTransacao(objNFsPag As ClassNFsPag, objContabil As ClassContabil) As Long
'Exclui a Nota Fiscal com código interno passado por parâmetro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFPag_Exclui_EmTransacao
      
    lErro = Rotina_ObterObjetoFuncao("CPR_NFPAG_EXCLUI_EMTRANSACAO", X)
    If lErro <> SUCESSO Then Error 65190
    
    NFPag_Exclui_EmTransacao = X.NFPag_Exclui_EmTransacao(objNFsPag, objContabil)
    
    Exit Function
    
Erro_NFPag_Exclui_EmTransacao:

    NFPag_Exclui_EmTransacao = Err
    
    Select Case Err

        Case 65190
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145532)
        
    End Select
    
    Exit Function
    
End Function

Function TituloPagar_Exclui_EmTransacao(objTituloPagar As ClassTituloPagar, objContabil As ClassContabil) As Long
'Executa a exclusão do Título a Pagar Passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TituloPagar_Exclui_EmTransacao
      
    lErro = Rotina_ObterObjetoFuncao("CPR_TITULOPAGAR_EXCLUI_EMTRANSACAO", X)
    If lErro <> SUCESSO Then Error 65191
    
    TituloPagar_Exclui_EmTransacao = X.TituloPagar_Exclui_EmTransacao(objTituloPagar, objContabil)
    
    Exit Function
    
Erro_TituloPagar_Exclui_EmTransacao:

    TituloPagar_Exclui_EmTransacao = Err
    
    Select Case Err

        Case 65191
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145533)
        
    End Select
    
    Exit Function
    
End Function

Function CPR_Exclui(iClasseDocCPR As Integer, lNumIntDocCPR As Long, Optional iExclusao As Integer = 0) As Long
'Exclui o documento de CPR com a classe e o número interno passados
'Se iExclusao = 1 indica que se trata de uma exclusão fisica devido a exclusao de uma nota fiscal
'ATENÇÃO : ESSA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CPR_Exclui
      
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CPR_EXCLUI", X)
    If lErro <> SUCESSO Then Error 65194
    
    CPR_Exclui = X.CPR_Exclui(iClasseDocCPR, lNumIntDocCPR, iExclusao)
    
    Exit Function
    
Erro_CPR_Exclui:

    CPR_Exclui = Err
    
    Select Case Err

        Case 65194
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145534)
        
    End Select
    
    Exit Function
    
End Function


Function NFiscalNumAuto(objNFiscal As ClassNFiscal) As Long
'Lê o Proximo número na tabela de Série e Coloca no objNFiscal
'Faz Lock Exclusive e atualiza o Número na Tabela de Série

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalNumAuto
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCALNUMAUTO", X)
    If lErro <> SUCESSO Then gError 71183

    NFiscalNumAuto = X.NFiscalNumAuto(objNFiscal)

    Exit Function
    
Erro_NFiscalNumAuto:
    
    NFiscalNumAuto = gErr
    
    Select Case gErr

        Case 71183
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145535)
        
    End Select
    
    Exit Function

End Function



'VERSÃO 2

Function ComprasConfig_Le(objConfiguraCOM As ClassConfiguraCOM)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComprasConfig_Le
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRASCONFIG_LE", X)
    If lErro <> SUCESSO Then Error 54692
    
    ComprasConfig_Le = X.ComprasConfig_Le(objConfiguraCOM)

    Exit Function
    
Erro_ComprasConfig_Le:
    
    ComprasConfig_Le = Err
    
    Select Case Err

        Case 54692
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145536)
        
    End Select
    
    Exit Function

End Function

Function ColComprasConfig_MoverCampos_ConfiguraCOM(objConfiguraCOM As ClassConfiguraCOM, colComprasConfig As colComprasConfig)
'Preenche objConfiguraCOM com os dados de objComprasConfig

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ColComprasConfig_MoverCampos_ConfiguraCOM
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_COLCOMPRASCONFIG_MOVERCAMPOS_CONFCOM", X)
    If lErro <> SUCESSO Then Error 54693
    
    ColComprasConfig_MoverCampos_ConfiguraCOM = X.ColComprasConfig_MoverCampos_ConfiguraCOM(objConfiguraCOM, colComprasConfig)

    Exit Function
    
Erro_ColComprasConfig_MoverCampos_ConfiguraCOM:
    
    ColComprasConfig_MoverCampos_ConfiguraCOM = Err
    
    Select Case Err

        Case 54693
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145537)
        
    End Select
    
    Exit Function

End Function

Function ConfiguraCOM_MoverCampos_ColComprasConfig(objConfiguraCOM As ClassConfiguraCOM, colComprasConfig As colComprasConfig) As Long
'Transmite os dados do objConfiguraCOM para ColConfiguraCOM

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ConfiguraCOM_MoverCampos_ColComprasConfig
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONFIGURACOM_MOVERCAMPOS_COLCOMPRASC", X)
    If lErro <> SUCESSO Then Error 54694
    
    ConfiguraCOM_MoverCampos_ColComprasConfig = X.ConfiguraCOM_MoverCampos_ColComprasConfig(objConfiguraCOM, colComprasConfig)

    Exit Function
    
Erro_ConfiguraCOM_MoverCampos_ColComprasConfig:
    
    ConfiguraCOM_MoverCampos_ColComprasConfig = Err
    
    Select Case Err

        Case 54694
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145538)
        
    End Select
    
    Exit Function

End Function


Sub Converte_Status(ByVal iStatus As Integer, sStatus As String)
'converte o Status da Nota Fiscal do valor inteiro para o seu correspondente textual

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Converte_Status
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CONVERTE_STATUS", X)
    If lErro <> SUCESSO Then gError 74942
    
    
    Call X.Converte_Status(iStatus, sStatus)
    
    Exit Sub
    
Erro_Converte_Status:

    Select Case gErr

        Case 74942
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145539)
        
    End Select
    
    Exit Sub
    
End Sub

Function NFiscal_Le_1(objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Emissão

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_1
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE_1", X)
    If lErro <> SUCESSO Then gError 83972
    
    NFiscal_Le_1 = X.NFiscal_Le_1(objNFiscal)

    Exit Function
    
Erro_NFiscal_Le_1:
    
    NFiscal_Le_1 = gErr
    
    Select Case gErr

        Case 83972
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145540)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Ultima_Devolucao_Forn(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal de Devolução de Devolução da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Devolucao_Forn
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PESQUISA_ULTIMA_DEVOLUCAO_FORN", X)
    If lErro <> SUCESSO Then gError 89413
    
    Pesquisa_Ultima_Devolucao_Forn = X.Pesquisa_Ultima_Devolucao_Forn(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Ultima_Devolucao_Forn:
    
    Pesquisa_Ultima_Devolucao_Forn = gErr
    
    Select Case gErr

        Case 89413
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145541)
        
    End Select
    
    Exit Function
    
End Function

Function Pesquisa_Ultima_Devolucao_FornFil(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal de Devolução de Compra da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Devolucao_FornFil
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_PESQUISA_ULTIMA_DEVOLUCAO_FORNFIL", X)
    If lErro <> SUCESSO Then gError 89414
    
    Pesquisa_Ultima_Devolucao_FornFil = X.Pesquisa_Ultima_Devolucao_FornFil(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Ultima_Devolucao_FornFil:
    
    Pesquisa_Ultima_Devolucao_FornFil = gErr
    
    Select Case gErr

        Case 89414
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145542)
        
    End Select
    
    Exit Function
    
End Function

Function FilialFornecedor_Grava_Estatisticas(objNFiscal As ClassNFiscal, iCancelamento As Integer, Optional objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Tem que ser chamado dentro de Transação
'Atualiza a gravação da parte de estatística do Fornecedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialFornecedor_Grava_Estatisticas
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALFORNECEDOR_GRAVA_ESTATISTICA", X)
    If lErro <> SUCESSO Then Error 54670
    
    FilialFornecedor_Grava_Estatisticas = X.FilialFornecedor_Grava_Estatisticas(objNFiscal, iCancelamento, objFilFornFilEmp)

    Exit Function
    
Erro_FilialFornecedor_Grava_Estatisticas:
    
    FilialFornecedor_Grava_Estatisticas = Err
    
    Select Case Err

        Case 54670
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145543)
        
    End Select
    
    Exit Function

End Function

Public Function NFiscal_Le_SemNumIntDoc(objNFiscal As ClassNFiscal) As Long
'Lê os dados da nota fiscal a partir dos dados Numero, Serie, FilialEmpresa, Tipo, Fornecedor ou Cliente, FilialForn ou FilialCli, DataEmissao, DataEntrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_SemNumIntDoc
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_LE_SEMNUMINTDOC", X)
    If lErro <> SUCESSO Then gError 89669
    
    NFiscal_Le_SemNumIntDoc = X.NFiscal_Le_SemNumIntDoc(objNFiscal)

    Exit Function
    
Erro_NFiscal_Le_SemNumIntDoc:
    
    NFiscal_Le_SemNumIntDoc = gErr
    
    Select Case gErr

        Case 89669
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145544)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Grava_DataHora(objNFiscal As ClassNFiscal) As Long
'Atualiza os campos Data e Hora de entrada e saida da nota fiscal passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Grava_DataHora
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCAL_GRAVA_DATAHORA", X)
    If lErro <> SUCESSO Then gError 89720
    
    NFiscal_Grava_DataHora = X.NFiscal_Grava_DataHora(objNFiscal)

    Exit Function
    
Erro_NFiscal_Grava_DataHora:
    
    NFiscal_Grava_DataHora = gErr
    
    Select Case gErr

        Case 89720
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145545)
        
    End Select
    
    Exit Function

End Function

Function CPR_Altera(objNFiscal As ClassNFiscal) As Long
'Altera a data de emissao do titulo relacionado à nota fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CPR_Altera
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_CPR_ALTERA", X)
    If lErro <> SUCESSO Then gError 89732
    
    CPR_Altera = X.CPR_Altera(objNFiscal)

    Exit Function
    
Erro_CPR_Altera:
    
    CPR_Altera = gErr
    
    Select Case gErr

        Case 89732
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145546)
        
    End Select
    
    Exit Function

End Function

Function SldMesForn_Altera_Estatisticas(objNFiscal As ClassNFiscal, objNFiscalBD As ClassNFiscal) As Long
'Atualiza as tabelas SldMesForn e SldDiaForn diminuindo a quantidade e o valor da data de emissao original e colocando na nova data, se a data de emissão foi alterada
'Tem que ser chamado dentro de Transação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesForn_Altera_Estatisticas
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_SLDMESFORN_ALTERA_ESTATISTICAS", X)
    If lErro <> SUCESSO Then gError 89735
    
    SldMesForn_Altera_Estatisticas = X.SldMesForn_Altera_Estatisticas(objNFiscal, objNFiscalBD)

    Exit Function
    
Erro_SldMesForn_Altera_Estatisticas:
    
    SldMesForn_Altera_Estatisticas = gErr
    
    Select Case gErr

        Case 89735
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145547)
        
    End Select
    
    Exit Function

End Function

