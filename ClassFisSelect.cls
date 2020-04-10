VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFisSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86TiposArquivos_Le_Todos(ByVal colIN86TiposArquivo As Collection) As Long
'Lê todos os itens da tabela IN86TiposArquivos e carrega a coleção
'colIN86TiposArquivo RETORNA(Output) os tipos de arquivos lidos do BD

Dim lComando As Long
Dim lErro As Long
Dim tIN86TipoArquivo As typeIN86TiposArquivos
Dim objIN86TipoArquivo As ClassIN86TiposArquivos

On Error GoTo Erro_IN86TiposArquivos_Le_Todos

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103500

    'inicializa as strings
    tIN86TipoArquivo.sDescricao = String(STRING_IN86TIPOSARQUIVOS_DESCRICAO, 0)
    tIN86TipoArquivo.sPrefixoNome = String(STRING_IN86TIPOSARQUIVOS_PREFIXONOME, 0)
    tIN86TipoArquivo.sRotinaGeradora = String(STRING_IN86TIPOSARQUIVOS_ROTINAGERADORA, 0)
    tIN86TipoArquivo.sLayoutArquivo = String(STRING_IN86TIPOSARQUIVOS_LAYOUTARQUIVO, 0)

    'seleciona no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, PrefixoNome, RotinaGeradora, Auxiliar, LayoutArquivo FROM IN86TiposArquivos ORDER BY Codigo", _
    tIN86TipoArquivo.iCodigo, tIN86TipoArquivo.sDescricao, tIN86TipoArquivo.sPrefixoNome, tIN86TipoArquivo.sRotinaGeradora, tIN86TipoArquivo.iAuxiliar, tIN86TipoArquivo.sLayoutArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 103501

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103502

    'se deu sem dados -> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103503

    'Enquanto houver tipos de arquivos
    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia um objeto do tipo objIN86TipoArquivo
        Set objIN86TipoArquivo = New ClassIN86TiposArquivos

        'preenche seus atributos com os dados lidos do bd
        objIN86TipoArquivo.iCodigo = tIN86TipoArquivo.iCodigo
        objIN86TipoArquivo.sDescricao = tIN86TipoArquivo.sDescricao
        objIN86TipoArquivo.sPrefixoNome = tIN86TipoArquivo.sPrefixoNome
        objIN86TipoArquivo.sRotinaGeradora = tIN86TipoArquivo.sRotinaGeradora
        objIN86TipoArquivo.iAuxiliar = tIN86TipoArquivo.iAuxiliar
        objIN86TipoArquivo.sLayoutArquivo = tIN86TipoArquivo.sLayoutArquivo

        'adiciona-o à coleção, mantendo o codigo como chave para permitir acesso posterior
        colIN86TiposArquivo.Add objIN86TipoArquivo, CStr(objIN86TipoArquivo.iCodigo)

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103504

    Loop

    IN86TiposArquivos_Le_Todos = SUCESSO

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_IN86TiposArquivos_Le_Todos:

    IN86TiposArquivos_Le_Todos = gErr

    Select Case gErr

        Case 103500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103501, 103502, 103504
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86TIPOSARQUIVOS_LEITURA", gErr)

        Case 103503
        'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150408)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Modelos_Le_Todos(ByVal colIN86Modelos As Collection) As Long
'Le todos os modelos da tabela de IN86Modelos
'colIN86Modelos RETORNA(Output) os modelos lidos do BD

Dim lErro As Long
Dim objIN86Modelo As ClassIN86Modelos
Dim tIN86Modelo As typeIN86Modelos

Dim lComando As Long

On Error GoTo Erro_IN86Modelos_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103515

    'inicializa a string do tipo
    tIN86Modelo.sModelo = String(STRING_IN86MODELOS_MODELO, 0)

    'seleciona no bd
    lErro = Comando_Executar(lComando, "SELECT Codigo, Modelo, DataInicio, DataFim, MeioEntrega, Etiquetas  FROM IN86Modelos ORDER BY Codigo", _
    tIN86Modelo.iCodigo, tIN86Modelo.sModelo, tIN86Modelo.dtDataInicio, tIN86Modelo.dtDataFim, tIN86Modelo.iMeioEntrega, tIN86Modelo.iEtiquetas)
    If lErro <> AD_SQL_SUCESSO Then gError 103516

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103517

    'se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103519

    'enquanto houver registros no bd
    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia um objeto novo
        Set objIN86Modelo = New ClassIN86Modelos

        'preenche seus atributos
        objIN86Modelo.iCodigo = tIN86Modelo.iCodigo
        objIN86Modelo.sModelo = tIN86Modelo.sModelo
        objIN86Modelo.dtDataInicio = tIN86Modelo.dtDataInicio
        objIN86Modelo.dtDataFim = tIN86Modelo.dtDataFim
        objIN86Modelo.iMeioEntrega = tIN86Modelo.iMeioEntrega
        objIN86Modelo.iEtiquetas = tIN86Modelo.iEtiquetas

        'adiciona o item na coleção
        colIN86Modelos.Add objIN86Modelo

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103518

    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86Modelos_Le_Todos = SUCESSO

    Exit Function

Erro_IN86Modelos_Le_Todos:

    IN86Modelos_Le_Todos = gErr

    Select Case gErr

        Case 103515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103516, 103517, 103518
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_LEITURA", gErr)

        Case 103519
        'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150409)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86MeioEntrega_Le_Todos(ByVal colIN86MeiosEntrega As Collection) As Long
'Lê todos os meios de entrega da tabela IN86MeioEntrega
'colIN86MeiosEntrega RETORNA(Output) os meios de entrega lidos do BD

Dim lErro As Long
Dim lComando As Long
Dim objIN86MeioEntrega As ClassIN86MeioEntrega
Dim tIN86MeioEntrega As typeIN86MeioEntrega

On Error GoTo Erro_IN86MeioEntrega_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103508

    'inicializa a string do tipo
    tIN86MeioEntrega.sDescricao = String(STRING_IN86MEIOENTREGA_DESCRICAO, 0)

    'seleciona no bd todos os meios de entrega
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, UsaEtiqueta FROM IN86MeioEntrega ORDER BY Codigo", _
    tIN86MeioEntrega.iCodigo, tIN86MeioEntrega.sDescricao, tIN86MeioEntrega.iUsaEtiqueta)
    If lErro <> AD_SQL_SUCESSO Then gError 103509

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103510

    'se o bd estiver vazio--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103511

    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia o objIN86MeioEntrega
        Set objIN86MeioEntrega = New ClassIN86MeioEntrega

        'preenche os atributos do objeto
        objIN86MeioEntrega.iCodigo = tIN86MeioEntrega.iCodigo
        objIN86MeioEntrega.sDescricao = tIN86MeioEntrega.sDescricao
        objIN86MeioEntrega.iUsaEtiqueta = tIN86MeioEntrega.iUsaEtiqueta
        
        'adiciona o objeto à coleção
        colIN86MeiosEntrega.Add objIN86MeioEntrega

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103512

    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86MeioEntrega_Le_Todos = SUCESSO

    Exit Function

Erro_IN86MeioEntrega_Le_Todos:

    IN86MeioEntrega_Le_Todos = gErr

    Select Case gErr

        Case 103508
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103509, 103510, 103512
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MEIOENTREGA_LEITURA", gErr)

        Case 103511
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150410)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86MeioEntrega_Le(ByVal objIN86MeioEntrega As ClassIN86MeioEntrega) As Long
'Lê o meio de entrega com código passado como parâmetro
'objIN86MeioEntrega RECEBE(Input) o código do meio entrega que será lido
'e RETORNA(Output) os dados do meio de entrega

Dim lErro As Long
Dim tIN86MeioEntrega As typeIN86MeioEntrega
Dim lComando As Long

On Error GoTo Erro_IN86MeioEntrega_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103638
    
    'inicializa a String
    tIN86MeioEntrega.sDescricao = String(STRING_IN86MEIOENTREGA_DESCRICAO, 0)
    
    'Seleciona o meio de entrega no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, UsaEtiqueta FROM IN86MeioEntrega WHERE Codigo = ? ", tIN86MeioEntrega.iCodigo, tIN86MeioEntrega.sDescricao, tIN86MeioEntrega.iUsaEtiqueta, objIN86MeioEntrega.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103639
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103640
    
    If lErro = AD_SQL_SEM_DADOS Then gError 103641
    
    With tIN86MeioEntrega
        
        'preenche o objeto com os dados do meio de entrega lido
        objIN86MeioEntrega.iCodigo = .iCodigo
        objIN86MeioEntrega.sDescricao = .sDescricao
        objIN86MeioEntrega.iUsaEtiqueta = .iUsaEtiqueta
        
    End With
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86MeioEntrega_Le = SUCESSO

    Exit Function

Erro_IN86MeioEntrega_Le:

    IN86MeioEntrega_Le = gErr

    Select Case gErr

        Case 103638
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103639, 103640
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MEIOENTREGA_LEITURA", gErr)
        
        Case 103641
        'sera tratado na rotina chamadora
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150411)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Modelos_Le_Modelo(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Lê em IN86Modelos, o modelo com nome passado como parâmetro
'objIN86Modelo RECEBE(Input) o nome do modelo que será lido e
'RETORNA(Output) os dados do modelos

Dim lErro As Long
Dim lComando As Long
Dim tIN86Modelo As typeIN86Modelos

On Error GoTo Erro_IN86Modelos_Le_Modelo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103570

    'inicializa a String do tipo do modelo
    tIN86Modelo.sModelo = String(STRING_IN86MODELOS_MODELO, 0)
    
    'seleciona o modelo na tabela de modelos
    lErro = Comando_Executar(lComando, "SELECT Modelo, Codigo, DataInicio, DataFim, MeioEntrega, Etiquetas FROM IN86Modelos WHERE Modelo = ?", _
    tIN86Modelo.sModelo, tIN86Modelo.iCodigo, tIN86Modelo.dtDataInicio, tIN86Modelo.dtDataFim, tIN86Modelo.iMeioEntrega, tIN86Modelo.iEtiquetas, objIN86Modelo.sModelo)
    If lErro <> AD_SQL_SUCESSO Then gError 103571

    'busca o primeiro da seleção
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103572

    'se não existe-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103573

    'preenche o objeto com os dados do tipo
    objIN86Modelo.iCodigo = tIN86Modelo.iCodigo
    objIN86Modelo.sModelo = tIN86Modelo.sModelo
    objIN86Modelo.dtDataInicio = tIN86Modelo.dtDataInicio
    objIN86Modelo.dtDataFim = tIN86Modelo.dtDataFim
    objIN86Modelo.iMeioEntrega = tIN86Modelo.iMeioEntrega
    objIN86Modelo.iEtiquetas = tIN86Modelo.iEtiquetas

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86Modelos_Le_Modelo = SUCESSO

    Exit Function

Erro_IN86Modelos_Le_Modelo:

    IN86Modelos_Le_Modelo = gErr

    Select Case gErr

        Case 103570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103571, 103572
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_LEITURA", gErr)

        Case 103573
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150412)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Arquivos_Le_Modelo(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Lê os arquivos a partir do código do modelo passado como parâmetro
'objIN86Modelo RECEBE(Input) o código do modelo que terá seus arquivos lidos e
'RETORNA(Output) os arquivos lidos dentro de uma coleção

Dim lErro As Long
Dim tIN86Arquivo As typeIN86Arquivos
Dim objIN86Arquivo As ClassIN86Arquivos
Dim lComando As Long

On Error GoTo Erro_IN86Arquivos_Le_Modelo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103585

    'inicializar os buffers
    tIN86Arquivo.sNome = String(STRING_IN86ARQUIVOS_NOME, 0)

    'busca todos os arquivos referentes a um modelo específico
    lErro = Comando_Executar(lComando, "SELECT Tipo , Selecionado , Nome , RelatDUMP , RelatAcompanhamento , Layout , FilialEmpresa , NumEtiqueta FROM IN86Arquivos WHERE Modelo = ?", _
    tIN86Arquivo.iTipo, tIN86Arquivo.iSelecionado, tIN86Arquivo.sNome, tIN86Arquivo.iDUMP, tIN86Arquivo.iRelatAcompanhamento, tIN86Arquivo.iLayout, tIN86Arquivo.iFilialEmpresa, tIN86Arquivo.iNumEtiqueta, objIN86Modelo.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103586

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103587

    'se não achou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103588

    'enquanto houver arquivo para ser lido
    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia um objeto
        Set objIN86Arquivo = New ClassIN86Arquivos

        'preenche o obj com os dados lidos
        objIN86Arquivo.iModelo = tIN86Arquivo.iModelo
        objIN86Arquivo.iTipo = tIN86Arquivo.iTipo
        objIN86Arquivo.iSelecionado = tIN86Arquivo.iSelecionado
        objIN86Arquivo.sNome = tIN86Arquivo.sNome
        objIN86Arquivo.iDUMP = tIN86Arquivo.iDUMP
        objIN86Arquivo.iRelatAcompanhamento = tIN86Arquivo.iRelatAcompanhamento
        objIN86Arquivo.iLayout = tIN86Arquivo.iLayout
        objIN86Arquivo.iFilialEmpresa = tIN86Arquivo.iFilialEmpresa
        objIN86Arquivo.iNumEtiqueta = tIN86Arquivo.iNumEtiqueta
        
        'o adiciona à coleção de arquivos do modelo
        objIN86Modelo.colIN86Arquivos.Add objIN86Arquivo

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103589

    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86Arquivos_Le_Modelo = SUCESSO

    Exit Function

Erro_IN86Arquivos_Le_Modelo:

    IN86Arquivos_Le_Modelo = gErr

    Select Case gErr

        Case 103585
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103588 'não encontrado

        Case 103586, 103587, 103589
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_LEITURA", gErr, objIN86Modelo.sModelo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150413)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function Modelos_Le_Modelo_Completo(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Le um registro da tabela IN86Modelos e todos os seus relacionados
'na tabela IN86Arquivos e carrega o objin86Modelo
'objIN86Modelo RECEBE(Input) o código do modelo que será lido e
'RETORNA(Output) dados do modelo e os arquivos relacionados ao modelo

Dim lErro As Long

On Error GoTo Erro_Modelos_Le_Modelo_Completo

    'le um determinado modelo na tabela de IN86Moedelos
    lErro = CF("IN86Modelos_Le_Modelo", objIN86Modelo)
    If lErro <> SUCESSO And lErro <> 103573 Then gError 103590
    
    'se retornou SEM_DADOS
    If lErro = 103573 Then gError 103607

    'le todos os arquivos vinculados ao modelo lido
    lErro = CF("IN86Arquivos_Le_Modelo", objIN86Modelo)
    If lErro <> SUCESSO And lErro <> 103588 Then gError 103591
    
    'se retornou SEM_DADOS
    If lErro = 103588 Then gError 103608

    Modelos_Le_Modelo_Completo = SUCESSO

    Exit Function

Erro_Modelos_Le_Modelo_Completo:

    Modelos_Le_Modelo_Completo = gErr

    Select Case gErr

        Case 103590, 103591

        Case 103607, 103608 'tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150414)

    End Select
    
    Exit Function
    
End Function

Function FISConfig_Le(ByVal objFIS As ClassFIS) As Long

Dim lErro As Long
Dim sConteudo As String
Dim lComando As Long
Dim objFISConfig As New ClassFISConfig

On Error GoTo Erro_FISConfig_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80169

    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_CODIGO_FISCAL_SERVICOS"

    sConteudo = String(STRING_CONTEUDO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM FisConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, objFISConfig.sCodigo, objFISConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 80170

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80171

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then gError 80172
    
    objFIS.iCodFiscalServico = StrParaInt(sConteudo)
    
    'Libera comando
    Call Comando_Fechar(lComando)

    FISConfig_Le = SUCESSO

    Exit Function

Erro_FISConfig_Le:

    FISConfig_Le = gErr

    Select Case gErr

        Case 80169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80170, 80171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FISCONFIG", gErr)

        Case 80172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_FIS_CONFIG_NAO_ENCONTRADO", gErr, objFISConfig.sCodigo, objFISConfig.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150415)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Declan_Obter_Estoque(ByVal dtData As Date, ByVal iFilialEmpresa As Integer, dValorEstoque As Double, ByVal lComando As Long) As Long
'Le o valor de estoque na data passada para a Filial Passada

Dim lErro As Long
Dim tRegInventario As typeRegInventario

On Error GoTo Erro_Declan_Obter_Estoque
    
    If lComando = 0 Then
        'Abertura de Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 78070
    End If
    
    'Procura no BD o Registro de inventário com a data e Filial Empresa passados em objRegInventario
    lErro = Comando_Executar(lComando, "SELECT QuantidadeUMEstoque, ValorUnitario FROM RegInventario WHERE Data = ? AND FilialEmpresa = ?", tRegInventario.dQuantidadeUMEstoque, tRegInventario.dValorUnitario, dtData, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 78071

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78072

    Do While lErro = AD_SQL_SUCESSO

        'Acumula os valores em dValorEstoque
        dValorEstoque = tRegInventario.dQuantidadeUMEstoque * tRegInventario.dValorUnitario
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78073
    
    Loop
    
    Declan_Obter_Estoque = SUCESSO

    Exit Function

Erro_Declan_Obter_Estoque:

    Declan_Obter_Estoque = gErr

    Select Case gErr

        Case 78070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78071, 78072, 78073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150416)

    End Select

    Exit Function
    
End Function

Function LivRegES_Le_Saidas(dValorSaidasICMS As Double, dValorSaidasIPI As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, lComando As Long, Optional ByVal iTipo As Integer = LIVREGES_SAIDA_TODAS) As Long
'Lê o valor das Saidas de ICMS e IpI

Dim lErro As Long, dICMS As Double, dAcumICMS As Double, dIPI As Double, dAcumIPI As Double

On Error GoTo Erro_LivRegES_Le_Saidas

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 78036
    End If
    
    If iTipo = LIVREGES_SAIDA_EXTERNA Then
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ? AND LEFT(NaturezaOP,1) = '7'", dICMS, dIPI, TIPO_REGES_SAIDA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    ElseIf iTipo = LIVREGES_SAIDA_NACIONAL Then
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ? AND LEFT(NaturezaOP,1) <> '7'", dICMS, dIPI, TIPO_REGES_SAIDA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    Else
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ?", dICMS, dIPI, TIPO_REGES_SAIDA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 78037

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78038
    
    Do While lErro = AD_SQL_SUCESSO
    
        dAcumICMS = dAcumICMS + Arredonda_Moeda(dICMS, 2)
        dAcumIPI = dAcumIPI + Arredonda_Moeda(dIPI, 2)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78038
    
    Loop
    
    dValorSaidasICMS = dAcumICMS
    dValorSaidasIPI = dAcumIPI
    
    Exit Function
    
    LivRegES_Le_Saidas = SUCESSO
    
Erro_LivRegES_Le_Saidas:

    LivRegES_Le_Saidas = gErr
    
    Select Case gErr
            
        Case 78036
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78037, 78038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150417)

    End Select

    Exit Function
        
End Function

Function LivRegES_Le_Entradas(dValorEntradasICMS As Double, dValorEntradasIPI As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, lComando As Long, Optional ByVal iTipo As Integer = LIVREGES_ENTRADA_TODAS) As Long
'Lê o valor das Entradas de ICMS e IPI

Dim lErro As Long, dICMS As Double, dAcumICMS As Double, dIPI As Double, dAcumIPI As Double

On Error GoTo Erro_LivRegES_Le_Entradas

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 78041
    End If
    
    If iTipo = LIVREGES_ENTRADA_EXTERNA Then
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ? AND LEFT(NaturezaOP,1) = '3' ", dICMS, dIPI, TIPO_REGES_ENTRADA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    ElseIf iTipo = LIVREGES_ENTRADA_NACIONAL Then
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ? AND LEFT(NaturezaOP,1) <> '3'", dICMS, dIPI, TIPO_REGES_ENTRADA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    Else
        lErro = Comando_Executar(lComando, "SELECT ValorICMS, ValorIPI FROM LivRegES, LivRegESLinha WHERE LivRegES.NumIntDoc = LivRegESLinha.NumIntDocRegES AND Tipo = ? AND (CodigoICMS = ? OR ValorICMS <> 0 OR CodigoIPI = ? OR ValorIPI <> 0) AND FilialEmpresa = ? AND (StatusLivro = ? OR StatusLivro = ?) AND Data >=? AND Data <= ?", dICMS, dIPI, TIPO_REGES_ENTRADA, TIPO_TRIBUTADO, TIPO_TRIBUTADO, iFilialEmpresa, STATUS_FIS_ALTERADA, STATUS_FIS_ORIGINAL, dtDataDe, dtDataAte)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 78042
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78043
    
    Do While lErro = AD_SQL_SUCESSO
    
        dAcumICMS = dAcumICMS + Arredonda_Moeda(dICMS, 2)
        dAcumIPI = dAcumIPI + Arredonda_Moeda(dIPI, 2)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78043
    
    Loop
    
    dValorEntradasICMS = dAcumICMS
    dValorEntradasIPI = dAcumIPI
    
    LivRegES_Le_Entradas = SUCESSO
    
    Exit Function
    
Erro_LivRegES_Le_Entradas:

    LivRegES_Le_Entradas = gErr
    
    Select Case gErr
            
        Case 78041
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78042, 78043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150418)

    End Select

    Exit Function
        
End Function

Function ApuracaoICMS_Le_SaldoCredor(ByVal objApuracao As ClassRegApuracao, lComando As Long) As Long
'Le o Saldo credor do periodo anterior
'função de Relatorio

Dim lErro As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoICMS_Le_SaldoCredor

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 78046
    End If
    
    'Le o Saldo credor do periodo anterior
    lErro = Comando_Executar(lComando, "SELECT SaldoCredorInicial, SaldoCredorFinal FROM RegApuracaoICMS WHERE FilialEmpresa =? AND DataInicial = ? AND DataFinal = ?", tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, objApuracao.iFilialEmpresa, objApuracao.dtDataInicial, objApuracao.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 78047

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78048

    'Se não encontrou nenhuma apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 78049

    'Move os dados lidos para objApuracao
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    
    ApuracaoICMS_Le_SaldoCredor = SUCESSO

    Exit Function

Erro_ApuracaoICMS_Le_SaldoCredor:

    ApuracaoICMS_Le_SaldoCredor = gErr

    Select Case gErr

        Case 78046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78047, 78048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)

        Case 78049 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150419)

    End Select

    Exit Function

End Function

Function ApuracaoIPI_Le_SaldoCredor(ByVal objApuracao As ClassRegApuracao, lComando As Long) As Long
'Le o Saldo credor do periodo anterior
'função de Relatorio

Dim lErro As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoIPI_Le_SaldoCredor

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 78053
    End If
    
    'Le o Saldo credor do periodo anterior
    lErro = Comando_Executar(lComando, "SELECT SaldoCredorInicial, SaldoCredorFinal FROM RegApuracaoIPI WHERE FilialEmpresa =? AND DataInicial = ? AND DataFinal = ?", tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, objApuracao.iFilialEmpresa, objApuracao.dtDataInicial, objApuracao.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 78054

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78055

    'Move os dados lidos para objApuracao
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    
    ApuracaoIPI_Le_SaldoCredor = SUCESSO

    Exit Function

Erro_ApuracaoIPI_Le_SaldoCredor:

    ApuracaoIPI_Le_SaldoCredor = gErr

    Select Case gErr

        Case 78053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78054, 78055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150420)

    End Select

    Exit Function

End Function

Function LivrosFilial_Le(ByVal objLivroFilial As ClassLivrosFilial) As Long
'Lê dados do Livro do Tipo passado da Filial Empresa também passada

Dim lErro As Long
Dim tLivroFiscal As typeLivrosFilial
Dim lComando As Long

On Error GoTo Erro_LivrosFilial_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67989
    
    'Lê dados do Livro Fiscal da FilialEmpresa e do Tipo passados em objLivroFilial
    lErro = Comando_Executar(lComando, "SELECT CodLivro, FilialEmpresa, NumeroProxLivro, NumeroProxFolha, Imprime, Periodicidade, DataInicial, DataFinal, ImpressoEm FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", _
    tLivroFiscal.iCodLivro, tLivroFiscal.iFilialEmpresa, tLivroFiscal.iNumeroProxLivro, tLivroFiscal.iNumeroProxFolha, tLivroFiscal.iImprime, tLivroFiscal.iPeriodicidade, tLivroFiscal.dtDataInicial, tLivroFiscal.dtDataFinal, tLivroFiscal.dtImpressoEm, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67990
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67991
    
    'Se não encontrou o livro, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67992
        
    objLivroFilial.iCodLivro = tLivroFiscal.iCodLivro
    objLivroFilial.iFilialEmpresa = tLivroFiscal.iFilialEmpresa
    objLivroFilial.iNumeroProxLivro = tLivroFiscal.iNumeroProxLivro
    objLivroFilial.iNumeroProxFolha = tLivroFiscal.iNumeroProxFolha
    objLivroFilial.iImprime = tLivroFiscal.iImprime
    objLivroFilial.iPeriodicidade = tLivroFiscal.iPeriodicidade
    objLivroFilial.dtDataInicial = tLivroFiscal.dtDataInicial
    objLivroFilial.dtDataFinal = tLivroFiscal.dtDataFinal
    objLivroFilial.dtImpressoEm = tLivroFiscal.dtImpressoEm
        
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    LivrosFilial_Le = SUCESSO
    
    Exit Function

Erro_LivrosFilial_Le:
    
    LivrosFilial_Le = gErr
    
    Select Case gErr
    
        Case 67989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 67990, 67991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
        
        Case 67992 'Livro não cadastrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150421)
        
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ApuracaoICMS_Le_UltimaFechada(ByVal objApuracao As ClassRegApuracao) As Long
'Lê dados da última Apuração ICMS Fechada (ordem de data final)

Dim lErro As Long
Dim lComando As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoICMS_Le_UltimaFechada

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67976
    
    tApuracao.sBairro = String(STRING_BAIRRO_REGAPURACAO, 0)
    tApuracao.sCEP = String(STRING_CEP, 0)
    tApuracao.sCgc = String(STRING_CGC, 0)
    tApuracao.sComplemento = String(STRING_COMPLEMENTO, 0)
    tApuracao.sContato = String(STRING_CONTATO_REGAPURACAO, 0)
    tApuracao.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tApuracao.sLocalEntregaGIA = String(STRING_LOCALENTREGA, 0)
    tApuracao.sLogradouro = String(STRING_LOGRADOURO, 0)
    tApuracao.sMunicipio = String(STRING_MUNICIPIO, 0)
    tApuracao.sNome = String(STRING_FILIALEMPRESA_NOME, 0)
    tApuracao.sObservacoes = String(STRING_OBSERVACAO, 0)
    tApuracao.sTelContato = String(STRING_TELCONTATO, 0)
    tApuracao.sUF = String(STRING_ESTADO, 0)
    
    'Lê os dados das Apuração ICMS
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, DataInicial, DataFinal, NumIntDocLivFechado, SaldoCredorInicial, SaldoCredorFinal, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, DataImpressao FROM RegApuracaoICMS WHERE FilialEmpresa =? AND NumIntDocLivFechado <> 0 ORDER BY DataFinal DESC", _
    tApuracao.lNumIntDoc, tApuracao.iFilialEmpresa, tApuracao.dtDataInicial, tApuracao.dtDataFinal, tApuracao.lNumIntDocLivFechado, tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, tApuracao.dtDataEntregaGIA, tApuracao.sLocalEntregaGIA, tApuracao.sObservacoes, tApuracao.sCgc, tApuracao.sInscricaoEstadual, tApuracao.sNome, tApuracao.sMunicipio, tApuracao.sUF, tApuracao.sLogradouro, tApuracao.lNumero, tApuracao.sComplemento, tApuracao.sBairro, tApuracao.sCEP, tApuracao.sContato, tApuracao.sTelContato, tApuracao.dtDataImpressao, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67977

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67978

    'Se não encontrou nenhuma apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67979

    'Move os dados lidos para objApuracao
    objApuracao.lNumIntDoc = tApuracao.lNumIntDoc
    objApuracao.dtDataInicial = tApuracao.dtDataInicial
    objApuracao.dtDataFinal = tApuracao.dtDataFinal
    objApuracao.lNumIntDocLivFechado = tApuracao.lNumIntDocLivFechado
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    objApuracao.dtDataEntregaGIA = tApuracao.dtDataEntregaGIA
    objApuracao.sLocalEntregaGIA = tApuracao.sLocalEntregaGIA
    objApuracao.sObservacoes = tApuracao.sObservacoes
    objApuracao.sCgc = tApuracao.sCgc
    objApuracao.sInscricaoEstadual = tApuracao.sInscricaoEstadual
    objApuracao.sNome = tApuracao.sNome
    objApuracao.sMunicipio = tApuracao.sMunicipio
    objApuracao.sUF = tApuracao.sUF
    objApuracao.sLogradouro = tApuracao.sLogradouro
    objApuracao.lNumero = tApuracao.lNumero
    objApuracao.sComplemento = tApuracao.sComplemento
    objApuracao.sBairro = tApuracao.sBairro
    objApuracao.sCEP = tApuracao.sCEP
    objApuracao.sContato = tApuracao.sContato
    objApuracao.sTelContato = tApuracao.sTelContato
    objApuracao.dtDataImpressao = tApuracao.dtDataImpressao
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    ApuracaoICMS_Le_UltimaFechada = SUCESSO

    Exit Function

Erro_ApuracaoICMS_Le_UltimaFechada:

    ApuracaoICMS_Le_UltimaFechada = gErr

    Select Case gErr

        Case 67976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67977, 67978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)

        Case 67979 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150422)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ApuracaoICMS_Le(ByVal objApuracao As ClassRegApuracao) As Long
'Lê dados da Apuração ICMS a partir da data Inicial, data final e FilialEmpresa passados em objApuracao

Dim lErro As Long
Dim lComando As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoICMS_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70010
            
    tApuracao.sBairro = String(STRING_BAIRRO_REGAPURACAO, 0)
    tApuracao.sCEP = String(STRING_CEP, 0)
    tApuracao.sCgc = String(STRING_CGC, 0)
    tApuracao.sComplemento = String(STRING_COMPLEMENTO, 0)
    tApuracao.sContato = String(STRING_CONTATO_REGAPURACAO, 0)
    tApuracao.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tApuracao.sLocalEntregaGIA = String(STRING_LOCALENTREGA, 0)
    tApuracao.sLogradouro = String(STRING_LOGRADOURO, 0)
    tApuracao.sMunicipio = String(STRING_MUNICIPIO, 0)
    tApuracao.sNome = String(STRING_FILIALEMPRESA_NOME, 0)
    tApuracao.sObservacoes = String(STRING_OBSERVACAO, 0)
    tApuracao.sTelContato = String(STRING_TELCONTATO, 0)
    tApuracao.sUF = String(STRING_ESTADO, 0)
            
    'Lê dados da apuração ICMS a partir da data inicial, final e FilialEmpresa passados em objApuracao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, DataInicial, DataFinal, NumIntDocLivFechado, SaldoCredorInicial, SaldoCredorFinal, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, DataImpressao FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", _
    tApuracao.lNumIntDoc, tApuracao.iFilialEmpresa, tApuracao.dtDataInicial, tApuracao.dtDataFinal, tApuracao.lNumIntDocLivFechado, tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, tApuracao.dtDataEntregaGIA, tApuracao.sLocalEntregaGIA, tApuracao.sObservacoes, tApuracao.sCgc, tApuracao.sInscricaoEstadual, tApuracao.sNome, tApuracao.sMunicipio, tApuracao.sUF, tApuracao.sLogradouro, tApuracao.lNumero, tApuracao.sComplemento, tApuracao.sBairro, tApuracao.sCEP, tApuracao.sContato, tApuracao.sTelContato, tApuracao.dtDataImpressao, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 70011
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70012
    
    'Se não encontrou a Apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70013
    
    'Move os dados lidos para objApuracao
    objApuracao.lNumIntDoc = tApuracao.lNumIntDoc
    objApuracao.dtDataInicial = tApuracao.dtDataInicial
    objApuracao.dtDataFinal = tApuracao.dtDataFinal
    objApuracao.lNumIntDocLivFechado = tApuracao.lNumIntDocLivFechado
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    objApuracao.dtDataEntregaGIA = tApuracao.dtDataEntregaGIA
    objApuracao.sLocalEntregaGIA = tApuracao.sLocalEntregaGIA
    objApuracao.sObservacoes = tApuracao.sObservacoes
    objApuracao.sCgc = tApuracao.sCgc
    objApuracao.sInscricaoEstadual = tApuracao.sInscricaoEstadual
    objApuracao.sNome = tApuracao.sNome
    objApuracao.sMunicipio = tApuracao.sMunicipio
    objApuracao.sUF = tApuracao.sUF
    objApuracao.sLogradouro = tApuracao.sLogradouro
    objApuracao.lNumero = tApuracao.lNumero
    objApuracao.sComplemento = tApuracao.sComplemento
    objApuracao.sBairro = tApuracao.sBairro
    objApuracao.sCEP = tApuracao.sCEP
    objApuracao.sContato = tApuracao.sContato
    objApuracao.sTelContato = tApuracao.sTelContato
    objApuracao.dtDataImpressao = tApuracao.dtDataImpressao
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    ApuracaoICMS_Le = SUCESSO
    
    Exit Function
    
Erro_ApuracaoICMS_Le:
    
    ApuracaoICMS_Le = gErr
    
    Select Case gErr
            
        Case 70010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 70011, 70012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
        
        Case 70013 'Apuração ICMS não cadastrada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150423)
        
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RegApuracaoIPIItens_Le(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Lê o Item de Registro de Apuração IPI a partir do Tipo, da descrição e da data passados como parametros no objRegApuracaoItem

Dim lErro As Long
Dim lComando As Long
Dim tRegApuracaoItens As typeRegApuracaoItens

On Error GoTo Erro_RegApuracaoIPIItens_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79060

    tRegApuracaoItens.sDescricao = String(STRING_DESCRICAO_APURACAO, 0)

    'Lê os dados do Item de Registro de Apuração IPI a partir do Tipo, da descrição e da data
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocApuracao, TipoReg, Descricao, Data, Valor, FilialEmpresa FROM RegApuracaoIPIItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", tRegApuracaoItens.lNumIntDoc, tRegApuracaoItens.lNumIntDocApuracao, tRegApuracaoItens.iTipoReg, tRegApuracaoItens.sDescricao, tRegApuracaoItens.dtData, tRegApuracaoItens.dValor, tRegApuracaoItens.iFilialEmpresa, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79061

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79062

    'Se não encontrou o tipo de registro de apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79063

    'Passa dados para objRegApuracaoItem
    objRegApuracaoItem.lNumIntDoc = tRegApuracaoItens.lNumIntDoc
    objRegApuracaoItem.lNumIntDocApuracao = tRegApuracaoItens.lNumIntDocApuracao
    objRegApuracaoItem.iTipoReg = tRegApuracaoItens.iTipoReg
    objRegApuracaoItem.sDescricao = tRegApuracaoItens.sDescricao
    objRegApuracaoItem.dtData = tRegApuracaoItens.dtData
    objRegApuracaoItem.dValor = tRegApuracaoItens.dValor
    objRegApuracaoItem.iFilialEmpresa = tRegApuracaoItens.iFilialEmpresa
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    RegApuracaoIPIItens_Le = SUCESSO

    Exit Function

Erro_RegApuracaoIPIItens_Le:

    RegApuracaoIPIItens_Le = gErr

    Select Case gErr

        Case 79060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79061, 79062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)

        Case 79063 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150424)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoRegApuracaoIPI_Le(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Lê o tipo de apuração a partir do código passado

Dim lErro As Long
Dim lComando As Long
Dim tTipoRegApuracao As typeTiposRegApuracao

On Error GoTo Erro_TipoRegApuracaoIPI_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79021

    tTipoRegApuracao.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê o Tipo de apuração a partir do código
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Secao, PreCadastrado FROM TiposRegApuracaoIPI WHERE Codigo = ?", tTipoRegApuracao.iCodigo, tTipoRegApuracao.sDescricao, tTipoRegApuracao.iSecao, tTipoRegApuracao.iPreCadastrado, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79022

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79023

    'Se não encontrou o tipo de registro de apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79024

    'Passa dados para objTipoReg
    objTipoRegApuracao.iCodigo = tTipoRegApuracao.iCodigo
    objTipoRegApuracao.sDescricao = tTipoRegApuracao.sDescricao
    objTipoRegApuracao.iSecao = tTipoRegApuracao.iSecao
    objTipoRegApuracao.iPreCadastrado = tTipoRegApuracao.iPreCadastrado

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    TipoRegApuracaoIPI_Le = SUCESSO

    Exit Function

Erro_TipoRegApuracaoIPI_Le:

    TipoRegApuracaoIPI_Le = gErr

    Select Case gErr

        Case 79021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79022, 79023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOIPI", gErr)

        Case 79024 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150425)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RegApuracaoICMSItens_Le(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Lê o Item de Registro de Apuração ICMS a partir do Tipo, da descrição e da data passados como parametros no objRegApuracaoItem

Dim lErro As Long
Dim lComando As Long
Dim tRegApuracaoItens As typeRegApuracaoItens

On Error GoTo Erro_RegApuracaoICMSItens_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67939

    tRegApuracaoItens.sDescricao = String(STRING_DESCRICAO_APURACAO, 0)

    'Lê os dados do Item de Registro de Apuração ICMS a partir do Tipo, da descrição e da data
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocApuracao, TipoReg, Descricao, Data, Valor, FilialEmpresa FROM RegApuracaoICMSItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", tRegApuracaoItens.lNumIntDoc, tRegApuracaoItens.lNumIntDocApuracao, tRegApuracaoItens.iTipoReg, tRegApuracaoItens.sDescricao, tRegApuracaoItens.dtData, tRegApuracaoItens.dValor, tRegApuracaoItens.iFilialEmpresa, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67940

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67941

    'Se não encontrou o tipo de registro de apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67942

    'Passa dados para objRegApuracaoItem
    objRegApuracaoItem.lNumIntDoc = tRegApuracaoItens.lNumIntDoc
    objRegApuracaoItem.lNumIntDocApuracao = tRegApuracaoItens.lNumIntDocApuracao
    objRegApuracaoItem.iTipoReg = tRegApuracaoItens.iTipoReg
    objRegApuracaoItem.sDescricao = tRegApuracaoItens.sDescricao
    objRegApuracaoItem.dtData = tRegApuracaoItens.dtData
    objRegApuracaoItem.dValor = tRegApuracaoItens.dValor
    objRegApuracaoItem.iFilialEmpresa = tRegApuracaoItens.iFilialEmpresa
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    RegApuracaoICMSItens_Le = SUCESSO

    Exit Function

Erro_RegApuracaoICMSItens_Le:

    RegApuracaoICMSItens_Le = gErr

    Select Case gErr

        Case 67939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67940, 67941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)

        Case 67942 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150426)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoRegApuracaoICMS_Le(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Lê o tipo de apuração a partir do código passado

Dim lErro As Long
Dim lComando As Long
Dim tTipoRegApuracao As typeTiposRegApuracao

On Error GoTo Erro_TipoRegApuracaoICMS_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67890
    
    tTipoRegApuracao.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    
    'Lê o Tipo de apuração a partir do código
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Secao, PreCadastrado FROM TiposRegApuracaoICMS WHERE Codigo = ?", tTipoRegApuracao.iCodigo, tTipoRegApuracao.sDescricao, tTipoRegApuracao.iSecao, tTipoRegApuracao.iPreCadastrado, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 67891
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67892
    
    'Se não encontrou o tipo de registro de apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67893
    
    'Passa dados para objTipoReg
    objTipoRegApuracao.iCodigo = tTipoRegApuracao.iCodigo
    objTipoRegApuracao.sDescricao = tTipoRegApuracao.sDescricao
    objTipoRegApuracao.iSecao = tTipoRegApuracao.iSecao
    objTipoRegApuracao.iPreCadastrado = tTipoRegApuracao.iPreCadastrado
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    TipoRegApuracaoICMS_Le = SUCESSO
    
    Exit Function
    
Erro_TipoRegApuracaoICMS_Le:
    
    TipoRegApuracaoICMS_Le = gErr

    Select Case gErr
        
        Case 67890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67891, 67892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOICMS", gErr)
                    
        Case 67893 'Não cadastrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150427)

    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ApuracaoIPI_Le_UltimaFechada(ByVal objApuracao As ClassRegApuracao) As Long
'Lê dados da última Apuração IPI Fechada (ordem de data final)

Dim lErro As Long
Dim lComando As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoIPI_Le_UltimaFechada

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79112
    
    tApuracao.sBairro = String(STRING_BAIRRO_REGAPURACAO, 0)
    tApuracao.sCEP = String(STRING_CEP, 0)
    tApuracao.sCgc = String(STRING_CGC, 0)
    tApuracao.sComplemento = String(STRING_COMPLEMENTO, 0)
    tApuracao.sContato = String(STRING_CONTATO_REGAPURACAO, 0)
    tApuracao.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tApuracao.sLocalEntregaGIA = String(STRING_LOCALENTREGA, 0)
    tApuracao.sLogradouro = String(STRING_LOGRADOURO, 0)
    tApuracao.sMunicipio = String(STRING_MUNICIPIO, 0)
    tApuracao.sNome = String(STRING_FILIALEMPRESA_NOME, 0)
    tApuracao.sObservacoes = String(STRING_OBSERVACAO, 0)
    tApuracao.sTelContato = String(STRING_TELCONTATO, 0)
    tApuracao.sUF = String(STRING_ESTADO, 0)
    
    'Lê os dados das Apuração IPI
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, DataInicial, DataFinal, NumIntDocLivFechado, SaldoCredorInicial, SaldoCredorFinal, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, DataImpressao FROM RegApuracaoIPI WHERE FilialEmpresa =? AND NumIntDocLivFechado <> 0 ORDER BY DataFinal DESC", _
    tApuracao.lNumIntDoc, tApuracao.iFilialEmpresa, tApuracao.dtDataInicial, tApuracao.dtDataFinal, tApuracao.lNumIntDocLivFechado, tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, tApuracao.dtDataEntregaGIA, tApuracao.sLocalEntregaGIA, tApuracao.sObservacoes, tApuracao.sCgc, tApuracao.sInscricaoEstadual, tApuracao.sNome, tApuracao.sMunicipio, tApuracao.sUF, tApuracao.sLogradouro, tApuracao.lNumero, tApuracao.sComplemento, tApuracao.sBairro, tApuracao.sCEP, tApuracao.sContato, tApuracao.sTelContato, tApuracao.dtDataImpressao, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79113

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79114

    'Se não encontrou nenhuma apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79115

    'Move os dados lidos para objApuracao
    objApuracao.lNumIntDoc = tApuracao.lNumIntDoc
    objApuracao.dtDataInicial = tApuracao.dtDataInicial
    objApuracao.dtDataFinal = tApuracao.dtDataFinal
    objApuracao.lNumIntDocLivFechado = tApuracao.lNumIntDocLivFechado
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    objApuracao.dtDataEntregaGIA = tApuracao.dtDataEntregaGIA
    objApuracao.sLocalEntregaGIA = tApuracao.sLocalEntregaGIA
    objApuracao.sObservacoes = tApuracao.sObservacoes
    objApuracao.sCgc = tApuracao.sCgc
    objApuracao.sInscricaoEstadual = tApuracao.sInscricaoEstadual
    objApuracao.sNome = tApuracao.sNome
    objApuracao.sMunicipio = tApuracao.sMunicipio
    objApuracao.sUF = tApuracao.sUF
    objApuracao.sLogradouro = tApuracao.sLogradouro
    objApuracao.lNumero = tApuracao.lNumero
    objApuracao.sComplemento = tApuracao.sComplemento
    objApuracao.sBairro = tApuracao.sBairro
    objApuracao.sCEP = tApuracao.sCEP
    objApuracao.sContato = tApuracao.sContato
    objApuracao.sTelContato = tApuracao.sTelContato
    objApuracao.dtDataImpressao = tApuracao.dtDataImpressao
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    ApuracaoIPI_Le_UltimaFechada = SUCESSO

    Exit Function

Erro_ApuracaoIPI_Le_UltimaFechada:

    ApuracaoIPI_Le_UltimaFechada = gErr

    Select Case gErr

        Case 79112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79113, 79114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)

        Case 79115 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150428)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ApuracaoIPI_Le(ByVal objApuracao As ClassRegApuracao) As Long
'Lê dados da Apuração IPI a partir da data Inicial, data final e FilialEmpresa passados em objApuracao

Dim lErro As Long
Dim lComando As Long
Dim tApuracao As typeRegApuracao

On Error GoTo Erro_ApuracaoIPI_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79116
            
    tApuracao.sBairro = String(STRING_BAIRRO_REGAPURACAO, 0)
    tApuracao.sCEP = String(STRING_CEP, 0)
    tApuracao.sCgc = String(STRING_CGC, 0)
    tApuracao.sComplemento = String(STRING_COMPLEMENTO, 0)
    tApuracao.sContato = String(STRING_CONTATO_REGAPURACAO, 0)
    tApuracao.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tApuracao.sLocalEntregaGIA = String(STRING_LOCALENTREGA, 0)
    tApuracao.sLogradouro = String(STRING_LOGRADOURO, 0)
    tApuracao.sMunicipio = String(STRING_MUNICIPIO, 0)
    tApuracao.sNome = String(STRING_FILIALEMPRESA_NOME, 0)
    tApuracao.sObservacoes = String(STRING_OBSERVACAO, 0)
    tApuracao.sTelContato = String(STRING_TELCONTATO, 0)
    tApuracao.sUF = String(STRING_ESTADO, 0)
            
    'Lê dados da apuração IPI a partir da data inicial, final e FilialEmpresa passados em objApuracao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, DataInicial, DataFinal, NumIntDocLivFechado, SaldoCredorInicial, SaldoCredorFinal, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, DataImpressao FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", _
    tApuracao.lNumIntDoc, tApuracao.iFilialEmpresa, tApuracao.dtDataInicial, tApuracao.dtDataFinal, tApuracao.lNumIntDocLivFechado, tApuracao.dSaldoCredorInicial, tApuracao.dSaldoCredorFinal, tApuracao.dtDataEntregaGIA, tApuracao.sLocalEntregaGIA, tApuracao.sObservacoes, tApuracao.sCgc, tApuracao.sInscricaoEstadual, tApuracao.sNome, tApuracao.sMunicipio, tApuracao.sUF, tApuracao.sLogradouro, tApuracao.lNumero, tApuracao.sComplemento, tApuracao.sBairro, tApuracao.sCEP, tApuracao.sContato, tApuracao.sTelContato, tApuracao.dtDataImpressao, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 79117
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79118
    
    'Se não encontrou a Apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79119
    
    'Move os dados lidos para objApuracao
    objApuracao.lNumIntDoc = tApuracao.lNumIntDoc
    objApuracao.dtDataInicial = tApuracao.dtDataInicial
    objApuracao.dtDataFinal = tApuracao.dtDataFinal
    objApuracao.lNumIntDocLivFechado = tApuracao.lNumIntDocLivFechado
    objApuracao.dSaldoCredorInicial = tApuracao.dSaldoCredorInicial
    objApuracao.dSaldoCredorFinal = tApuracao.dSaldoCredorFinal
    objApuracao.dtDataEntregaGIA = tApuracao.dtDataEntregaGIA
    objApuracao.sLocalEntregaGIA = tApuracao.sLocalEntregaGIA
    objApuracao.sObservacoes = tApuracao.sObservacoes
    objApuracao.sCgc = tApuracao.sCgc
    objApuracao.sInscricaoEstadual = tApuracao.sInscricaoEstadual
    objApuracao.sNome = tApuracao.sNome
    objApuracao.sMunicipio = tApuracao.sMunicipio
    objApuracao.sUF = tApuracao.sUF
    objApuracao.sLogradouro = tApuracao.sLogradouro
    objApuracao.lNumero = tApuracao.lNumero
    objApuracao.sComplemento = tApuracao.sComplemento
    objApuracao.sBairro = tApuracao.sBairro
    objApuracao.sCEP = tApuracao.sCEP
    objApuracao.sContato = tApuracao.sContato
    objApuracao.sTelContato = tApuracao.sTelContato
    objApuracao.dtDataImpressao = tApuracao.dtDataImpressao
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    ApuracaoIPI_Le = SUCESSO
    
    Exit Function
    
Erro_ApuracaoIPI_Le:
    
    ApuracaoIPI_Le = gErr
    
    Select Case gErr
            
        Case 79116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 79117, 79118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
        
        Case 79119 'Apuração IPI não cadastrada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150429)
        
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function GNRICMS_Le(ByVal objGNRICMS As ClassGNRICMS) As Long
'Lê dados da Guia de Recolhimento de ICMS a partir do Código passado

Dim lErro As Long
Dim lComando As Long
Dim tGNRCICMS As typeGNRICMS

On Error GoTo Erro_GNRICMS_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70067
    
    tGNRCICMS.sCGCSubstTrib = String(STRING_CGC, 0)
    tGNRCICMS.sConvenio = String(STRING_CONVENIO, 0)
    tGNRCICMS.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tGNRCICMS.sNumero = String(STRING_NUMERO, 0)
    tGNRCICMS.sUFDestino = String(STRING_ESTADO, 0)
    tGNRCICMS.sUFSubstTrib = String(STRING_ESTADO, 0)
    tGNRCICMS.sCodReceita = String(STRING_GNRICMS_CODRECEITA, 0)
    tGNRCICMS.sCodObrigRecolher = String(STRING_GNRICMS_CODOBRIGRECOLHER, 0)
        
    'Lê dados do Guia de ICMS a partir do código passado em objGNRICMS
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, Tipo, DataPagto, CGCSubstTrib, InscricaoEstadual, UFSubstTrib, UFDestino, Banco, Agencia, Numero, Valor, Vencimento, DataRef, Convenio, NumIntRegApuracaoICMS, NumIntArqICMS, CodObrigRecolher, CodReceita FROM GNRICMS WHERE Codigo = ?", _
    tGNRCICMS.lNumIntDoc, tGNRCICMS.lCodigo, tGNRCICMS.iTipo, tGNRCICMS.dtDataPagto, tGNRCICMS.sCGCSubstTrib, tGNRCICMS.sInscricaoEstadual, tGNRCICMS.sUFSubstTrib, tGNRCICMS.sUFDestino, tGNRCICMS.iBanco, tGNRCICMS.iAgencia, tGNRCICMS.sNumero, tGNRCICMS.dValor, tGNRCICMS.dtDataVencimento, tGNRCICMS.dtDataRef, tGNRCICMS.sConvenio, tGNRCICMS.lNumIntRegApuracaoICMS, tGNRCICMS.lNumIntArqICMS, tGNRCICMS.sCodObrigRecolher, tGNRCICMS.sCodReceita, objGNRICMS.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 70068

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70069

    'Se não encontrou a Guia de ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70070

    'Preenche objGNRICMS com os dados lidos do BD
    objGNRICMS.lNumIntDoc = tGNRCICMS.lNumIntDoc
    objGNRICMS.lCodigo = tGNRCICMS.lCodigo
    objGNRICMS.iTipo = tGNRCICMS.iTipo
    objGNRICMS.dtDataPagto = tGNRCICMS.dtDataPagto
    objGNRICMS.sCGCSubstTrib = tGNRCICMS.sCGCSubstTrib
    objGNRICMS.sInscricaoEstadual = tGNRCICMS.sInscricaoEstadual
    objGNRICMS.sUFSubstTrib = tGNRCICMS.sUFSubstTrib
    objGNRICMS.sUFDestino = tGNRCICMS.sUFDestino
    objGNRICMS.iBanco = tGNRCICMS.iBanco
    objGNRICMS.iAgencia = tGNRCICMS.iAgencia
    objGNRICMS.sNumero = tGNRCICMS.sNumero
    objGNRICMS.dValor = tGNRCICMS.dValor
    objGNRICMS.dtVencimento = tGNRCICMS.dtDataVencimento
    objGNRICMS.dtDataRef = tGNRCICMS.dtDataRef
    objGNRICMS.sConvenio = tGNRCICMS.sConvenio
    objGNRICMS.lNumIntRegApuracaoICMS = tGNRCICMS.lNumIntRegApuracaoICMS
    objGNRICMS.lNumIntArqICMS = tGNRCICMS.lNumIntArqICMS
    objGNRICMS.sCodReceita = tGNRCICMS.sCodReceita
    objGNRICMS.sCodObrigRecolher = tGNRCICMS.sCodObrigRecolher
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    GNRICMS_Le = SUCESSO

    Exit Function

Erro_GNRICMS_Le:

    GNRICMS_Le = gErr

    Select Case gErr

        Case 70067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70068, 70069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GNRICMS", gErr)

        Case 70070 'Guia ICMS não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150430)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InfoArqICMS_Le(ByVal objInfoArqICMS As ClassInfoArqICMS) As Long
'Le registro em infoarqicms da filialempresa, com as datas Iniciais e Finais passadas

Dim lErro As Long
Dim tInfoArqICMS As typeInfoArqICMS
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_InfoArqICMS_Le

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69973
    Next
    
    tInfoArqICMS.sNomeArquivo = String(STRING_NOME_ARQ_COMPLETO, 0)
    tInfoArqICMS.sCgc = String(STRING_CGC, 0)
    tInfoArqICMS.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tInfoArqICMS.sNome = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tInfoArqICMS.sMunicipio = String(STRING_MUNICIPIO, 0)
    'tInfoArqICMS.sBairro = String(STRING_BAIRRO, 0) Comentado por Luiz: a string já está sendo inicializada mais abaixo e STRING_BAIRRO não existe mais
    tInfoArqICMS.sUF = String(STRING_ESTADO, 0)
    tInfoArqICMS.sLogradouro = String(STRING_LOGRADOURO, 0)
    tInfoArqICMS.sComplemento = String(STRING_COMPLEMENTO, 0)
    tInfoArqICMS.sBairro = String(STRING_BAIRRO_REGAPURACAO, 0)
    tInfoArqICMS.sCEP = String(STRING_CEP, 0)
    tInfoArqICMS.sContato = String(STRING_CONTATO_REGAPURACAO, 0)
    tInfoArqICMS.sTelContato = String(STRING_TELCONTATO, 0)

    'Verifica se existe um Arquivo de ICMS com as datas limites e filialempresa passados
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, DataInicial, DataFinal, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, NomeArquivo FROM InfoArqICMS WHERE FilialEmpresa = ? AND DataInicial = ? AND DataFinal = ?", tInfoArqICMS.lNumIntDoc, tInfoArqICMS.dtDataInicial, tInfoArqICMS.dtDataFinal, tInfoArqICMS.sCgc, tInfoArqICMS.sInscricaoEstadual, tInfoArqICMS.sNome, tInfoArqICMS.sMunicipio, tInfoArqICMS.sUF, tInfoArqICMS.sLogradouro, tInfoArqICMS.lNumero, tInfoArqICMS.sComplemento, tInfoArqICMS.sBairro, tInfoArqICMS.sCEP, tInfoArqICMS.sContato, tInfoArqICMS.sTelContato, tInfoArqICMS.sNomeArquivo, objInfoArqICMS.iFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 69974

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69975

    If lErro = AD_SQL_SEM_DADOS Then gError 69976
        
    objInfoArqICMS.lNumIntDoc = tInfoArqICMS.lNumIntDoc
    objInfoArqICMS.dtDataInicial = tInfoArqICMS.dtDataInicial
    objInfoArqICMS.dtDataFinal = tInfoArqICMS.dtDataFinal
    objInfoArqICMS.sCgc = tInfoArqICMS.sCgc
    objInfoArqICMS.sInscricaoEstadual = tInfoArqICMS.sInscricaoEstadual
    objInfoArqICMS.sNome = tInfoArqICMS.sNome
    objInfoArqICMS.sMunicipio = tInfoArqICMS.sMunicipio
    objInfoArqICMS.sUF = tInfoArqICMS.sUF
    objInfoArqICMS.sLogradouro = tInfoArqICMS.sLogradouro
    objInfoArqICMS.lNumero = tInfoArqICMS.lNumero
    objInfoArqICMS.sComplemento = tInfoArqICMS.sComplemento
    objInfoArqICMS.sBairro = tInfoArqICMS.sBairro
    objInfoArqICMS.sCEP = tInfoArqICMS.sCEP
    objInfoArqICMS.sContato = tInfoArqICMS.sContato
    objInfoArqICMS.sTelContato = tInfoArqICMS.sTelContato
    objInfoArqICMS.sNomeArquivo = tInfoArqICMS.sNomeArquivo
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    InfoArqICMS_Le = SUCESSO

    Exit Function

Erro_InfoArqICMS_Le:

    InfoArqICMS_Le = gErr

    Select Case gErr

        Case 69973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69974, 69975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOARQICMS", gErr)

        Case 69976
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150431)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Tributos_Le(ByVal colTributos As Collection) As Long
'Lê Tributos que possuem Livro Fiscal

Dim lErro As Long
Dim lComando As Long
Dim tTributo As typeTributo
Dim objTributo As ClassTributos

On Error GoTo Erro_Tributos_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70143

    tTributo.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê os Tributos que possuem Livro Fiscal (Livro = 1)
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, ApuracaoPeriodicidade FROM Tributos WHERE Livro = ?", tTributo.iCodigo, tTributo.sDescricao, tTributo.iApuracaoPeriodicidade, TRIBUTO_COM_LIVROFISCAL)
    If lErro <> AD_SQL_SUCESSO Then gError 70144

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70145

    'Se não encontrou nenhum Tributo, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70146

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO
                
        Set objTributo = New ClassTributos
        
        'Passa dados para objTributo
        objTributo.iCodigo = tTributo.iCodigo
        objTributo.sDescricao = tTributo.sDescricao
        objTributo.iApuracaoPeriodicidade = tTributo.iApuracaoPeriodicidade
        
        'Adiciona na coleção de Tributos
        colTributos.Add objTributo
        
        'Busca Próximo Tributo com Livro Fiscal
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70147
        
    Loop
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Tributos_Le = SUCESSO

    Exit Function

Erro_Tributos_Le:

    Tributos_Le = gErr

    Select Case gErr

        Case 70143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70144, 70145, 70147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTOS", gErr)

        Case 70146 'Não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150432)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Periodicidades_Le(ByVal colPeriodo As Collection) As Long
'Lê as Periodicidades dos Livros Fiscais

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sDescricao As String
Dim objPeriodo As AdmCodigoNome

On Error GoTo Erro_Periodicidades_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70151

    sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê as periodicidades da tabela PeriodicidadesLivrosFisc
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM PeriodicidadesLivrosFisc", iCodigo, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 70152

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70153

    'Se não encontrou nenhuma periodicidade, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70154

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO
                
        Set objPeriodo = New AdmCodigoNome
        
        'Passa dados para objPeriodo
        objPeriodo.iCodigo = iCodigo
        objPeriodo.sNome = sDescricao
        
        'Adiciona na coleção de Periodicidades
        colPeriodo.Add objPeriodo
        
        'Busca Próximo Periodicidade
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70155
        
    Loop
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Periodicidades_Le = SUCESSO

    Exit Function

Erro_Periodicidades_Le:

    Periodicidades_Le = gErr

    Select Case gErr

        Case 70151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70152, 70153, 70155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODICIDADESLIVROSFISC", gErr)

        Case 70154 'Não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150433)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LivrosFiscais_Le(ByVal iTributo As Integer, ByVal colLivroFiscal As Collection) As Long
'Lê Livros Fiscais associados ao Tributo passado por iTributo

Dim lErro As Long
Dim lComando As Long
Dim tLivroFiscal As typeLivroFiscal
Dim objLivroFiscal As ClassLivrosFiscais

On Error GoTo Erro_LivrosFiscais_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70169

    tLivroFiscal.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê os Livros Fiscais com o Tributo passado
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Periodicidade FROM LivrosFiscais WHERE CodTributo = ?", tLivroFiscal.iCodigo, tLivroFiscal.sDescricao, tLivroFiscal.iPeriodicidade, iTributo)
    If lErro <> AD_SQL_SUCESSO Then gError 70170

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70171

    'Se não encontrou nenhum Livro Fiscal, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70172

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO
                
        Set objLivroFiscal = New ClassLivrosFiscais
        
        'Passa dados para objLivroFiscal
        objLivroFiscal.iCodigo = tLivroFiscal.iCodigo
        objLivroFiscal.sDescricao = tLivroFiscal.sDescricao
        objLivroFiscal.iPeriodicidade = tLivroFiscal.iPeriodicidade
        
        'Adiciona na coleção de Livros Fiscais
        colLivroFiscal.Add objLivroFiscal
        
        'Busca Próximo Livro Fiscal
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70173
        
    Loop
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    LivrosFiscais_Le = SUCESSO

    Exit Function

Erro_LivrosFiscais_Le:

    LivrosFiscais_Le = gErr

    Select Case gErr

        Case 70169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70170, 70171, 70173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROFISCAL", gErr)

        Case 70172 'Não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150434)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LivroFiscal_Le_Codigo(ByVal objLivroFiscal As ClassLivrosFiscais) As Long
'Lê dados do Livro Fiscal a partir do código passado

Dim lErro As Long
Dim lComando As Long
Dim tLivroFiscal As typeLivroFiscal

On Error GoTo Erro_LivroFiscal_Le_Codigo

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70206

    tLivroFiscal.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê os Livros Fiscais com o código passado
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, CodTributo, Periodicidade, Apuracao FROM LivrosFiscais WHERE Codigo = ?", tLivroFiscal.iCodigo, tLivroFiscal.sDescricao, tLivroFiscal.iCodTributo, tLivroFiscal.iPeriodicidade, tLivroFiscal.iApuracao, objLivroFiscal.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 70207

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70208

    'Se não encontrou o Livro Fiscal, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70209
            
    'Passa dados para objLivroFiscal
    objLivroFiscal.iCodigo = tLivroFiscal.iCodigo
    objLivroFiscal.sDescricao = tLivroFiscal.sDescricao
    objLivroFiscal.iPeriodicidade = tLivroFiscal.iPeriodicidade
    objLivroFiscal.iApuracao = tLivroFiscal.iApuracao
    objLivroFiscal.iCodTributo = tLivroFiscal.iCodTributo
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    LivroFiscal_Le_Codigo = SUCESSO
    
    Exit Function

Erro_LivroFiscal_Le_Codigo:
    
    LivroFiscal_Le_Codigo = gErr
    
    Select Case gErr
    
        Case 70206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 70207, 70208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROFISCAL", gErr)
        
        Case 70209 'Livro não cadastrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150435)
        
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function LivrosFechados_Le_UltimaData(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Lê dados do último Livro Fiscal Fechado do Livro de código e Filial Empresea passados em objLivroFechado

Dim lErro As Long
Dim lComando As Long
Dim tLivroFechado As typeLivroFechado

On Error GoTo Erro_LivrosFechados_Le_UltimaData

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70228
        
    'Lê os dados do Livro Fiscal do código passado pro ordem decrescente da Data
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, CodLivro, FilialEmpresa, NumeroLivro, DataInicial, DataFinal, DataImpressao, FolhaInicial, FolhaFinal FROM LivrosFechados WHERE CodLivro = ? AND FilialEmpresa = ? ORDER BY DataFinal DESC", _
    tLivroFechado.lNumIntDoc, tLivroFechado.iCodLivro, tLivroFechado.iFilialEmpresa, tLivroFechado.iNumeroLivro, tLivroFechado.dtDataInicial, tLivroFechado.dtDataFinal, tLivroFechado.dtDataImpressao, tLivroFechado.iFolhaInicial, tLivroFechado.iFolhaFinal, objLivroFechado.iCodLivro, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70229

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70230

    'Se não encontrou nenhuma Livro Fechado, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70231

    'Move os dados lidos para objLivroFechado
    objLivroFechado.lNumIntDoc = tLivroFechado.lNumIntDoc
    objLivroFechado.iCodLivro = tLivroFechado.iCodLivro
    objLivroFechado.iFilialEmpresa = tLivroFechado.iFilialEmpresa
    objLivroFechado.iNumeroLivro = tLivroFechado.iNumeroLivro
    objLivroFechado.dtDataInicial = tLivroFechado.dtDataInicial
    objLivroFechado.dtDataFinal = tLivroFechado.dtDataFinal
    objLivroFechado.dtDataImpressao = tLivroFechado.dtDataImpressao
    objLivroFechado.iFolhaInicial = tLivroFechado.iFolhaInicial
    objLivroFechado.iFolhaFinal = tLivroFechado.iFolhaFinal
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    LivrosFechados_Le_UltimaData = SUCESSO

    Exit Function

Erro_LivrosFechados_Le_UltimaData:

    LivrosFechados_Le_UltimaData = gErr

    Select Case gErr

        Case 70228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70229, 70230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)

        Case 70231 'Não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150436)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LivRegES_Le_NotaFiscal(ByVal objLivRegES As ClassLivRegES) As Long
'Le o Livro correspondente a Nota Fiscal passada
'chave lNumIntDoc de Nota Fiscal

Dim lErro As Long
Dim tLivRegES As typeLivRegES
Dim lComando As Long

On Error GoTo Erro_LivRegES_Le_NotaFiscal
    
    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69461
    
    tLivRegES.sCgc = String(STRING_CGC, 0)
    tLivRegES.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tLivRegES.sNome = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tLivRegES.sUF = String(STRING_ESTADO, 0)
    tLivRegES.sSerie = String(STRING_SERIE, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocLivFechado, TipoNumIntDocOrigem, NumIntDocOrigem, StatusLivro, Tipo, CGC, InscricaoEstadual, Nome, Data, UF, Serie, NumNotaFiscal, Situacao, Modelo, Emitente, Destinatario, Origem, NumIntEmitente, NumIntRegApuracaoICMS, NumIntRegApuracaoIPI FROM LivRegES WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND FilialEmpresa = ?", tLivRegES.lNumIntDoc, tLivRegES.lNumIntLivroFechado, tLivRegES.iTipoNumIntDocOrigem, tLivRegES.lNumIntDocOrigem, tLivRegES.iStatusLivro, tLivRegES.iTipo, tLivRegES.sCgc, tLivRegES.sInscricaoEstadual, tLivRegES.sNome, tLivRegES.dtData, tLivRegES.sUF, tLivRegES.sSerie, tLivRegES.lNumNotaFiscal, tLivRegES.iSituacao, tLivRegES.iModelo, tLivRegES.iEmitente, tLivRegES.iDestinatario, tLivRegES.iOrigem, tLivRegES.lNumIntEmitente, tLivRegES.lNumIntRegApuracaoICMS, tLivRegES.lNumIntRegApuracaoIPI, TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL, objLivRegES.lNumIntNF, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69462

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69463
    
    If lErro <> AD_SQL_SUCESSO Then gError 69464
    
    objLivRegES.lNumIntDoc = tLivRegES.lNumIntDoc
    objLivRegES.lNumIntDocLivFechado = tLivRegES.lNumIntLivroFechado
    objLivRegES.iTipoNumIntDocOrigem = tLivRegES.iTipoNumIntDocOrigem
    objLivRegES.lNumIntDocOrigem = tLivRegES.lNumIntDocOrigem
    objLivRegES.iStatusLivro = tLivRegES.iStatusLivro
    objLivRegES.iTipo = tLivRegES.iTipo
    objLivRegES.sCgc = tLivRegES.sCgc
    objLivRegES.sInscricaoEstadual = tLivRegES.sInscricaoEstadual
    objLivRegES.sNome = tLivRegES.sNome
    objLivRegES.dtData = tLivRegES.dtData
    objLivRegES.sUF = tLivRegES.sUF
    objLivRegES.sSerie = tLivRegES.sSerie
    objLivRegES.lNumNotaFiscal = tLivRegES.lNumNotaFiscal
    objLivRegES.iSituacao = tLivRegES.iSituacao
    objLivRegES.iModelo = tLivRegES.iModelo
    objLivRegES.iEmitente = tLivRegES.iEmitente
    objLivRegES.iDestinatario = tLivRegES.iDestinatario
    objLivRegES.iOrigem = tLivRegES.iOrigem
    objLivRegES.lNumIntEmitente = tLivRegES.lNumIntEmitente
    objLivRegES.lNumIntRegApuracaoICMS = tLivRegES.lNumIntRegApuracaoICMS
    objLivRegES.lNumIntRegApuracaoIPI = tLivRegES.lNumIntRegApuracaoIPI
    
    'Le as linhas do livro para a Nota Fiscal passada
    lErro = CF("LivRegESLinha_Le", objLivRegES)
    If lErro <> SUCESSO Then gError 69465
    
    'Le os Itens de Notafiscal do livro para a Nota Fiscal passada
    lErro = CF("LivRegESItemNF_Le", objLivRegES)
    If lErro <> SUCESSO Then gError 69466
    
    Call Comando_Fechar(lComando)
    
    LivRegES_Le_NotaFiscal = SUCESSO
    
    Exit Function
    
Erro_LivRegES_Le_NotaFiscal:
    
    LivRegES_Le_NotaFiscal = gErr
    
    Select Case gErr
        
        Case 69461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69462, 69463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 69464 'Não encontrado
        
        Case 69465, 69466
                
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150437)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

Function LivRegESLinha_Le(ByVal objLivRegES As ClassLivRegES) As Long
'Le as Linhas para o Livro passado em objLivRegES.lNumIntDoc

Dim lErro As Long
Dim tLivRegESLinha As typeLivRegESLinha
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim lComando As Long

On Error GoTo Erro_LivRegESLinha_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69467
    
    tLivRegESLinha.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tLivRegESLinha.sClassifContabil = String(STRING_CONTA, 0)
    tLivRegESLinha.sObservacaoLivFisc = String(STRING_OBSERVACAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocRegES, NaturezaOp, ValorTotal, ValorICMSBase, ValorICMS, ValorICMSIsentoNaoTrib, ValorICMSOutras, ValorICMSSubstBase, ValorICMSSubstRet, AliquotaICMS, ValorDespAcess, ValorIPI, AliquotaIPI, ValorIPIBase, ValorIPIIsentoNaoTrib, ValorIPIOutras, ClassifContabil, ValorContabil, ObservacaoLivFisc, CodigoICMS, CodigoIPI FROM LivRegESLinha WHERE NumIntDocRegES = ? ORDER BY NumIntDoc", tLivRegESLinha.lNumIntDoc, tLivRegESLinha.lNumIntDocRegES, tLivRegESLinha.sNaturezaOp, tLivRegESLinha.dValorTotal, tLivRegESLinha.dValorICMSBase, tLivRegESLinha.dValorICMS, tLivRegESLinha.dValorICMSIsentoNaoTrib, tLivRegESLinha.dValorICMSOutras, tLivRegESLinha.dValorICMSSubstBase, tLivRegESLinha.dValorICMSSubstRet, tLivRegESLinha.dAliquotaICMS, tLivRegESLinha.dValorDespAcess, tLivRegESLinha.dValorIPI, tLivRegESLinha.dAliquotaIPI, tLivRegESLinha.dValorIPIBase, tLivRegESLinha.dValorIPIIsentoNaoTrib _
        , tLivRegESLinha.dValorIPIOutras, tLivRegESLinha.sClassifContabil, tLivRegESLinha.dValorContabil, tLivRegESLinha.sObservacaoLivFisc, tLivRegESLinha.iCodigoICMS, tLivRegESLinha.iCodigoIPI, objLivRegES.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69468

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69469
    
    Set objLivRegES.colLinhas = New Collection
    
    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLivRegESLinha = New ClassLivRegESLinha
        
        objLivRegESLinha.lNumIntDoc = tLivRegESLinha.lNumIntDoc
        objLivRegESLinha.lNumIntDocRegES = tLivRegESLinha.lNumIntDocRegES
        objLivRegESLinha.sNaturezaOp = tLivRegESLinha.sNaturezaOp
        objLivRegESLinha.dValorTotal = tLivRegESLinha.dValorTotal
        objLivRegESLinha.dValorICMSBase = tLivRegESLinha.dValorICMSBase
        objLivRegESLinha.dValorICMS = tLivRegESLinha.dValorICMS
        objLivRegESLinha.dValorICMSIsentoNaoTrib = tLivRegESLinha.dValorICMSIsentoNaoTrib
        objLivRegESLinha.dValorICMSOutras = tLivRegESLinha.dValorICMSOutras
        objLivRegESLinha.dValorICMSSubstBase = tLivRegESLinha.dValorICMSSubstBase
        objLivRegESLinha.dValorICMSSubstRet = tLivRegESLinha.dValorICMSSubstRet
        objLivRegESLinha.dAliquotaICMS = tLivRegESLinha.dAliquotaICMS
        objLivRegESLinha.dValorDespAcess = tLivRegESLinha.dValorDespAcess
        objLivRegESLinha.dValorIPI = tLivRegESLinha.dValorIPI
        objLivRegESLinha.dAliquotaIPI = tLivRegESLinha.dAliquotaIPI
        objLivRegESLinha.dValorIPIBase = tLivRegESLinha.dValorIPIBase
        objLivRegESLinha.dValorIPIIsentoNaoTrib = tLivRegESLinha.dValorIPIIsentoNaoTrib
        objLivRegESLinha.dValorIPIOutras = tLivRegESLinha.dValorIPIOutras
        objLivRegESLinha.sClassifContabil = tLivRegESLinha.sClassifContabil
        objLivRegESLinha.dValorContabil = tLivRegESLinha.dValorContabil
        objLivRegESLinha.sObservacaoLivFisc = tLivRegESLinha.sObservacaoLivFisc
        objLivRegESLinha.iCodigoICMS = tLivRegESLinha.iCodigoICMS
        objLivRegESLinha.iCodigoIPI = tLivRegESLinha.iCodigoIPI
        
        objLivRegES.colLinhas.Add objLivRegESLinha
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69470
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    LivRegESLinha_Le = SUCESSO
    
    Exit Function
    
Erro_LivRegESLinha_Le:

    LivRegESLinha_Le = gErr
    
    Select Case gErr
    
        Case 69467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69468, 69469, 69470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150438)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LivRegESItemNF_Le(ByVal objLivRegES As ClassLivRegES) As Long
'Le os Itens para o Livro passado em objLivRegES.lNumIntDoc

Dim lErro As Long
Dim tLivRegESItemNF As typeLivRegESItemNF
Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim lComando As Long

On Error GoTo Erro_LivRegESItemNF_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69471
    
    tLivRegESItemNF.sCFOP = String(STRING_NATUREZAOP_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocRegES, NumItem, CFOP, NumIntCadProd, Quantidade, ValorProduto, ValorDescontoDespAcess, ValorICMSBase, ValorICMSSubstBase, ValorIPI, AliquotaICMS, TipoTribICMS, TipoTribIPI, ValorICMS, ValorIPIBase, AliquotaIPI, RedBaseICMS, RedBaseIPI, TipoTributacao, AliquotaSubst, ValorSubst FROM LivRegESItemNF WHERE NumIntDocRegES = ? ORDER BY NumIntDoc", tLivRegESItemNF.lNumIntDoc, tLivRegESItemNF.lNumIntDocRegES, tLivRegESItemNF.iNumItem, tLivRegESItemNF.sCFOP, tLivRegESItemNF.lNumIntCadProd, tLivRegESItemNF.dQuantidade, tLivRegESItemNF.dValorProduto, tLivRegESItemNF.dValorDescontoDespAcess, tLivRegESItemNF.dValorICMSBase, tLivRegESItemNF.dValorICMSSubstBase, tLivRegESItemNF.dValorIPI, tLivRegESItemNF.dAliquotaICMS, tLivRegESItemNF.iTipoTribICMS, tLivRegESItemNF.iTipoTribIPI, tLivRegESItemNF.dValorICMS, tLivRegESItemNF.dValorIPIBase, tLivRegESItemNF.dAliquotaIPI _
    , tLivRegESItemNF.dRedBaseICMS, tLivRegESItemNF.dRedBaseIPI, tLivRegESItemNF.iTipoTributacao, tLivRegESItemNF.dAliquotaSubst, tLivRegESItemNF.dValorSubst, objLivRegES.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69472

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69473
    
    Set objLivRegES.ColItensNF = New Collection
    
    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLivRegESItemNF = New ClassLivRegESItemNF
        
        objLivRegESItemNF.lNumIntDoc = tLivRegESItemNF.lNumIntDoc
        objLivRegESItemNF.lNumIntDocRegES = tLivRegESItemNF.lNumIntDocRegES
        objLivRegESItemNF.iNumItem = tLivRegESItemNF.iNumItem
        objLivRegESItemNF.sCFOP = tLivRegESItemNF.sCFOP
        objLivRegESItemNF.lNumIntCadProd = tLivRegESItemNF.lNumIntCadProd
        objLivRegESItemNF.dQuantidade = tLivRegESItemNF.dQuantidade
        objLivRegESItemNF.dValorProduto = tLivRegESItemNF.dValorProduto
        objLivRegESItemNF.dValorDescontoDespAcess = tLivRegESItemNF.dValorDescontoDespAcess
        objLivRegESItemNF.dValorICMSBase = tLivRegESItemNF.dValorICMSBase
        objLivRegESItemNF.dValorICMSSubstBase = tLivRegESItemNF.dValorICMSSubstBase
        objLivRegESItemNF.dValorIPI = tLivRegESItemNF.dValorIPI
        objLivRegESItemNF.dAliquotaICMS = tLivRegESItemNF.dAliquotaICMS
        objLivRegESItemNF.iTipoTribICMS = tLivRegESItemNF.iTipoTribICMS
        objLivRegESItemNF.iTipoTribIPI = tLivRegESItemNF.iTipoTribIPI
        objLivRegESItemNF.dValorICMS = tLivRegESItemNF.dValorICMS
        objLivRegESItemNF.dValorIPIBase = tLivRegESItemNF.dValorIPIBase
        objLivRegESItemNF.dAliquotaIPI = tLivRegESItemNF.dAliquotaIPI
        objLivRegESItemNF.dRedBaseICMS = tLivRegESItemNF.dRedBaseICMS
        objLivRegESItemNF.dRedBaseIPI = tLivRegESItemNF.dRedBaseIPI
        objLivRegESItemNF.iTipoTributacao = tLivRegESItemNF.iTipoTributacao
        objLivRegESItemNF.dAliquotaSubst = tLivRegESItemNF.dAliquotaSubst
        objLivRegESItemNF.dValorSubst = tLivRegESItemNF.dValorSubst
    
        objLivRegES.ColItensNF.Add objLivRegESItemNF
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69474
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    LivRegESItemNF_Le = SUCESSO
    
    Exit Function
    
Erro_LivRegESItemNF_Le:

    LivRegESItemNF_Le = gErr
    
    Select Case gErr
    
        Case 69471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69472, 69473, 69474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150439)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RegInventario_Le_Data(ByVal objRegInventario As ClassRegInventario) As Long
'Lê o Registro de Inventário a partir da Data e Filial Empresa passados

Dim lErro As Long
Dim lComando As Long
Dim tRegInventario As typeRegInventario

On Error GoTo Erro_RegInventario_Le_Data

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70234

    tRegInventario.sContaContabil = String(STRING_CONTA, 0)
    tRegInventario.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tRegInventario.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tRegInventario.sModelo = String(STRING_PRODUTO_MODELO, 0)
    tRegInventario.sObservacoes = String(STRING_OBSERVACAO, 0)
    tRegInventario.sProduto = String(STRING_PRODUTO, 0)
    tRegInventario.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Procura no BD o Registro de inventário com a data e Filial Empresa passados em objRegInventario
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Produto, Data, NumIntDocLivFechado, Descricao, Modelo, IPICodigo, SiglaUMEstoque, QuantidadeUMEstoque, ValorUnitario, Natureza, QtdeNossaEmTerc, QtdeDeTercConosco, Observacoes, ContaContabil FROM RegInventario WHERE Data = ? AND FilialEmpresa = ?", _
    tRegInventario.iFilialEmpresa, tRegInventario.sProduto, tRegInventario.dtData, tRegInventario.lNumIntDocLivFechado, tRegInventario.sDescricao, tRegInventario.sModelo, tRegInventario.sIPICodigo, tRegInventario.sSiglaUMEstoque, tRegInventario.dQuantidadeUMEstoque, tRegInventario.dValorUnitario, tRegInventario.iNatureza, tRegInventario.dQtdeNossaEmTerc, tRegInventario.dQtdeDeTercConosco, tRegInventario.sObservacoes, tRegInventario.sContaContabil, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70235

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70236

    'Se não encontrou o Registro de inventário, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70237

    'Passa dados lidos para o objRegInventario
    objRegInventario.iFilialEmpresa = tRegInventario.iFilialEmpresa
    objRegInventario.sProduto = tRegInventario.sProduto
    objRegInventario.dtData = tRegInventario.dtData
    objRegInventario.lNumIntDocLivFechado = tRegInventario.lNumIntDocLivFechado
    objRegInventario.sDescricao = tRegInventario.sDescricao
    objRegInventario.sModelo = tRegInventario.sModelo
    objRegInventario.sIPICodigo = tRegInventario.sIPICodigo
    objRegInventario.sSiglaUMEstoque = tRegInventario.sSiglaUMEstoque
    objRegInventario.dQuantidadeUMEstoque = tRegInventario.dQuantidadeUMEstoque
    objRegInventario.dValorUnitario = tRegInventario.dValorUnitario
    objRegInventario.iNatureza = tRegInventario.iNatureza
    objRegInventario.dQtdeNossaEmTerc = tRegInventario.dQtdeNossaEmTerc
    objRegInventario.dQtdeDeTercConosco = tRegInventario.dQtdeDeTercConosco
    objRegInventario.sObservacoes = tRegInventario.sObservacoes
    objRegInventario.sContaContabil = tRegInventario.sContaContabil

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    RegInventario_Le_Data = SUCESSO

    Exit Function

Erro_RegInventario_Le_Data:

    RegInventario_Le_Data = gErr

    Select Case gErr

        Case 70234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70235, 70236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case 70237 'Registro de Inventário não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150440)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RegInventario_Le(ByVal objRegInventario As ClassRegInventario) As Long
'Lê o Registro de Inventário a partir da Data, Produto e Filial Empresa passados

Dim lErro As Long
Dim lComando As Long
Dim tRegInventario As typeRegInventario

On Error GoTo Erro_RegInventario_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70305

    tRegInventario.sContaContabil = String(STRING_CONTA, 0)
    tRegInventario.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tRegInventario.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tRegInventario.sModelo = String(STRING_PRODUTO_MODELO, 0)
    tRegInventario.sObservacoes = String(STRING_OBSERVACAO, 0)
    tRegInventario.sProduto = String(STRING_PRODUTO, 0)
    tRegInventario.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Procura no BD o Registro de inventário com o Produto, a data e Filial Empresa passados em objRegInventario
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Produto, Data, NumIntDocLivFechado, Descricao, Modelo, IPICodigo, SiglaUMEstoque, QuantidadeUMEstoque, ValorUnitario, Natureza, QtdeNossaEmTerc, QtdeDeTercConosco, Observacoes, ContaContabil, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef, QuantBenef3, CustoConsig3, CustoConsig, CustoDemo3, CustoDemo, CustoConserto3, CustoConserto, CustoOutras3, CustoOutras, CustoBenef, CustoBenef3 FROM RegInventario WHERE Produto = ? AND Data = ? AND FilialEmpresa = ?", _
    tRegInventario.iFilialEmpresa, tRegInventario.sProduto, tRegInventario.dtData, tRegInventario.lNumIntDocLivFechado, tRegInventario.sDescricao, tRegInventario.sModelo, tRegInventario.sIPICodigo, tRegInventario.sSiglaUMEstoque, tRegInventario.dQuantidadeUMEstoque, tRegInventario.dValorUnitario, tRegInventario.iNatureza, tRegInventario.dQtdeNossaEmTerc, tRegInventario.dQtdeDeTercConosco, tRegInventario.sObservacoes, tRegInventario.sContaContabil, tRegInventario.dQuantConsig3, tRegInventario.dQuantConsig, tRegInventario.dQuantDemo3, tRegInventario.dQuantDemo, tRegInventario.dQuantConserto3, tRegInventario.dQuantConserto, tRegInventario.dQuantOutras3, tRegInventario.dQuantOutras, tRegInventario.dQuantBenef, tRegInventario.dQuantBenef3, tRegInventario.dCustoConsig3, tRegInventario.dCustoConsig, _
    tRegInventario.dCustoDemo3, tRegInventario.dCustoDemo, tRegInventario.dCustoConserto3, tRegInventario.dCustoConserto, tRegInventario.dCustoOutras3, tRegInventario.dCustoOutras, tRegInventario.dCustoBenef, tRegInventario.dCustoBenef3, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70306

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70307

    'Se não encontrou o Registro de inventário, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70308

    'Passa dados lidos para o objRegInventario
    objRegInventario.iFilialEmpresa = tRegInventario.iFilialEmpresa
    objRegInventario.sProduto = tRegInventario.sProduto
    objRegInventario.dtData = tRegInventario.dtData
    objRegInventario.lNumIntDocLivFechado = tRegInventario.lNumIntDocLivFechado
    objRegInventario.sDescricao = tRegInventario.sDescricao
    objRegInventario.sModelo = tRegInventario.sModelo
    objRegInventario.sIPICodigo = tRegInventario.sIPICodigo
    objRegInventario.sSiglaUMEstoque = tRegInventario.sSiglaUMEstoque
    objRegInventario.dQuantidadeUMEstoque = tRegInventario.dQuantidadeUMEstoque
    objRegInventario.dValorUnitario = tRegInventario.dValorUnitario
    objRegInventario.iNatureza = tRegInventario.iNatureza
    objRegInventario.dQtdeNossaEmTerc = tRegInventario.dQtdeNossaEmTerc
    objRegInventario.dQtdeDeTercConosco = tRegInventario.dQtdeDeTercConosco
    objRegInventario.sObservacoes = tRegInventario.sObservacoes
    objRegInventario.sContaContabil = tRegInventario.sContaContabil
    objRegInventario.dQuantConsig3 = tRegInventario.dQuantConsig3
    objRegInventario.dQuantConsig = tRegInventario.dQuantConsig
    objRegInventario.dQuantDemo3 = tRegInventario.dQuantDemo3
    objRegInventario.dQuantDemo = tRegInventario.dQuantDemo
    objRegInventario.dQuantConserto3 = tRegInventario.dQuantConserto3
    objRegInventario.dQuantConserto = tRegInventario.dQuantConserto
    objRegInventario.dQuantOutras3 = tRegInventario.dQuantOutras3
    objRegInventario.dQuantOutras = tRegInventario.dQuantOutras
    objRegInventario.dQuantBenef = tRegInventario.dQuantBenef
    objRegInventario.dQuantBenef3 = tRegInventario.dQuantBenef3
    objRegInventario.dCustoConsig3 = tRegInventario.dCustoConsig3
    objRegInventario.dCustoConsig = tRegInventario.dCustoConsig
    objRegInventario.dCustoDemo3 = tRegInventario.dCustoDemo3
    objRegInventario.dCustoDemo = tRegInventario.dCustoDemo
    objRegInventario.dCustoConserto3 = tRegInventario.dCustoConserto3
    objRegInventario.dCustoConserto = tRegInventario.dCustoConserto
    objRegInventario.dCustoOutras3 = tRegInventario.dCustoOutras3
    objRegInventario.dCustoOutras = tRegInventario.dCustoOutras
    objRegInventario.dCustoBenef = tRegInventario.dCustoBenef
    objRegInventario.dCustoBenef3 = tRegInventario.dCustoBenef3
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    RegInventario_Le = SUCESSO

    Exit Function

Erro_RegInventario_Le:

    RegInventario_Le = gErr

    Select Case gErr

        Case 70305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70306, 70307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case 70308 'Registro de Inventário não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150441)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RegInventarioAlmox_Le(ByVal objRegInventarioAlmox As ClassRegInventarioAlmox) As Long
'Lê o Registro de Inventário a partir da Data, Produto e Filial Empresa passados

Dim lErro As Long
Dim lComando As Long
Dim tRegInventarioAlmox As typeRegInventarioAlmox

On Error GoTo Erro_RegInventarioAlmox_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69882

    tRegInventarioAlmox.sProduto = String(STRING_PRODUTO, 0)

    'Procura no BD o Registro de inventário com o Produto, a data e Filial Empresa passados em objRegInventario
    lErro = Comando_Executar(lComando, "SELECT Produto, Data, Almoxarifado, QuantidadeUMEstoque, QtdeDeTercConosco, QuantConsig3, QuantDemo3, QuantConserto3, QuantOutras3, QuantBenef3 FROM RegInventarioAlmox WHERE Produto = ? AND Data = ? AND Almoxarifado = ?", _
    tRegInventarioAlmox.sProduto, tRegInventarioAlmox.dtData, tRegInventarioAlmox.iAlmoxarifado, tRegInventarioAlmox.dQuantidadeUMEstoque, tRegInventarioAlmox.dQtdeDeTercConosco, tRegInventarioAlmox.dQuantConsig3, tRegInventarioAlmox.dQuantDemo3, tRegInventarioAlmox.dQuantConserto3, tRegInventarioAlmox.dQuantOutras3, tRegInventarioAlmox.dQuantBenef3, objRegInventarioAlmox.sProduto, objRegInventarioAlmox.dtData, objRegInventarioAlmox.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 69883

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69884

    'Se não encontrou o Registro de inventário, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 69885

    'Passa dados lidos para o objRegInventario
    objRegInventarioAlmox.sProduto = tRegInventarioAlmox.sProduto
    objRegInventarioAlmox.dtData = tRegInventarioAlmox.dtData
    objRegInventarioAlmox.iAlmoxarifado = tRegInventarioAlmox.iAlmoxarifado
    objRegInventarioAlmox.dQuantidadeUMEstoque = tRegInventarioAlmox.dQuantidadeUMEstoque
    objRegInventarioAlmox.dQtdeDeTercConosco = tRegInventarioAlmox.dQtdeDeTercConosco
    objRegInventarioAlmox.dQuantConsig3 = tRegInventarioAlmox.dQuantConsig3
    objRegInventarioAlmox.dQuantDemo3 = tRegInventarioAlmox.dQuantDemo3
    objRegInventarioAlmox.dQuantConserto3 = tRegInventarioAlmox.dQuantConserto3
    objRegInventarioAlmox.dQuantOutras3 = tRegInventarioAlmox.dQuantOutras3
    objRegInventarioAlmox.dQuantBenef3 = tRegInventarioAlmox.dQuantBenef3
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    RegInventarioAlmox_Le = SUCESSO

    Exit Function

Erro_RegInventarioAlmox_Le:

    RegInventarioAlmox_Le = gErr

    Select Case gErr

        Case 69882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69883, 69884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIOALMOX", gErr)

        Case 69885 'Registro de Inventário Almoxarifado não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150442)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RegInventario_IntervaloData_Critica(ByVal objLivroFilial As ClassLivrosFilial, ByVal dtData As Date) As Long
'Verifica a data passada está no intervalo de datas do Livro Fiscal de Registro de inventário aberto

Dim lErro As Long
Dim lComando As Long
Dim iCodLivro As Integer

On Error GoTo Erro_RegInventario_IntervaloData_Critica

    'Abertura de comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70922

    'Procura por Livro Fiscal aberto de Registro de inventário com a data passada dentro do intervalo de data inicial e final
    lErro = Comando_Executar(lComando, "SELECT CodLivro FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ? AND DataInicial <= ? AND DataFinal >=?", iCodLivro, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 70923

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70924

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70925

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RegInventario_IntervaloData_Critica = SUCESSO

    Exit Function

Erro_RegInventario_IntervaloData_Critica:

    RegInventario_IntervaloData_Critica = gErr

    Select Case gErr

        Case 70922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70923, 70924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 70925 'Não cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150443)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RegInventario_Obter_Quantidades(ByVal objRegInventario As ClassRegInventario) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69890
    Next
    
    If Day(objRegInventario.dtData) = 31 And Month(objRegInventario.dtData) = 12 Then
        
        'Obtem as quantidades nossa
        lErro = RegInventario_Obter_Quantidades_SldMesEstIni(objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 81678

        'Obtem quantidade de Terceiro conosco
        lErro = RegInventario_Obter_Quantidades_SldMesEst1Ini(objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 81679
    
        'Obtem quantidade de nosso em terceiros
        lErro = RegInventario_Obter_Quantidades_SldMesEst2Ini(objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 81680
    
    Else
    
        'Obtem as quantidades nossa
        lErro = CF("RegInventario_Obter_Quantidades_SldMesEst", objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 69904
        
        'Obtem quantidade de Terceiro conosco
        lErro = CF("RegInventario_Obter_Quantidades_SldMesEst1", objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 69905
    
        'Obtem quantidade de nosso em terceiros
        lErro = CF("RegInventario_Obter_Quantidades_SldMesEst2", objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 69906
    
        'O que houve de movimento neste mês
        lErro = CF("RegInventario_Obter_Quant_Mov_UltimoMes", objRegInventario, alComando())
        If lErro <> SUCESSO Then gError 69907

    End If
    
    'Obtem a quantidade em processo
    lErro = RegInventario_Obter_Quant_EmProcesso(objRegInventario, alComando())
    If lErro <> SUCESSO Then gError 69907

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RegInventario_Obter_Quantidades = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades:

    RegInventario_Obter_Quantidades = gErr

    Select Case gErr
                
        Case 69890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69904, 69905, 69906, 69907, 81678, 81679, 81680
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150444)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario
'??? poderia otimizar p/ler apenas os campos necessarios
Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim iMes As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim tProduto As typeProduto
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)
    
    '----------------------------- Leitura de SldMesEst ------------------
    'Lê as quantidades Inciais, Entrada e Saída para o Produto
    sComandoSQL = "SELECT QuantInicial, ValorInicial, QuantInicialCusto, ValorInicialCusto "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", " & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice) & ", " & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", " & "SaldoQuantCusto" & CStr(iIndice) & ", " & "SaldoValorCusto" & CStr(iIndice)
    Next
    
    sComandoSQL = sComandoSQL & ", QuantInicialRecIndl, ValorInicialRecIndl, SaldoQuantRecIndl1, SaldoQuantRecIndl2, SaldoQuantRecIndl3, SaldoQuantRecIndl4, SaldoQuantRecIndl5, SaldoQuantRecIndl6, SaldoQuantRecIndl7, SaldoQuantRecIndl8, SaldoQuantRecIndl9, SaldoQuantRecIndl10, SaldoQuantRecIndl11, SaldoQuantRecIndl12, " & _
                    "SaldoValorRecIndl1, SaldoValorRecIndl2, SaldoValorRecIndl3, SaldoValorRecIndl4, SaldoValorRecIndl5, SaldoValorRecIndl6, SaldoValorRecIndl7, SaldoValorRecIndl8, SaldoValorRecIndl9, SaldoValorRecIndl10, SaldoValorRecIndl11, SaldoValorRecIndl12 "
        
    'Tabela, Filtro
    sComandoSQL = sComandoSQL & " FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adSaldoQuantCusto(1), .adSaldoValorCusto(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adSaldoQuantCusto(2), .adSaldoValorCusto(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adSaldoQuantCusto(3), .adSaldoValorCusto(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adSaldoQuantCusto(4), .adSaldoValorCusto(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adSaldoQuantCusto(5), .adSaldoValorCusto(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), .adSaldoQuantCusto(6), .adSaldoValorCusto(6), _
        .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adSaldoQuantCusto(7), .adSaldoValorCusto(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adSaldoQuantCusto(8), .adSaldoValorCusto(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adSaldoQuantCusto(9), .adSaldoValorCusto(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adSaldoQuantCusto(10), .adSaldoValorCusto(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adSaldoQuantCusto(11), .adSaldoValorCusto(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), .adSaldoQuantCusto(12), .adSaldoValorCusto(12), _
        .dQuantInicialRecIndl, .dValorInicialRecIndl, .adSaldoQuantRecIndl(1), .adSaldoQuantRecIndl(2), .adSaldoQuantRecIndl(3), .adSaldoQuantRecIndl(4), .adSaldoQuantRecIndl(5), .adSaldoQuantRecIndl(6), .adSaldoQuantRecIndl(7), .adSaldoQuantRecIndl(8), .adSaldoQuantRecIndl(9), .adSaldoQuantRecIndl(10), .adSaldoQuantRecIndl(11), .adSaldoQuantRecIndl(12), _
        .adSaldoValorRecIndl(1), .adSaldoValorRecIndl(2), .adSaldoValorRecIndl(3), .adSaldoValorRecIndl(4), .adSaldoValorRecIndl(5), .adSaldoValorRecIndl(6), .adSaldoValorRecIndl(7), .adSaldoValorRecIndl(8), .adSaldoValorRecIndl(9), .adSaldoValorRecIndl(10), .adSaldoValorRecIndl(11), .adSaldoValorRecIndl(12), _
        iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69891
    End With

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69892

    'Alteracao Dnaiel em 23/11/2001
    If lErro = AD_SQL_SEM_DADOS Then
        
        objRegInventario.dQuantidadeUMEstoque = 0
        objRegInventario.dSaldoQuantCusto = 0
        objRegInventario.dValorEstoque = 0
        objRegInventario.dSaldoValorCusto = 0
        
    Else
        'Acumula até o Mês - 1
        objRegInventario.dQuantidadeUMEstoque = tSldMesEst.dQuantInicial - tSldMesEst.dQuantInicialRecIndl
        objRegInventario.dSaldoQuantCusto = tSldMesEst.dQuantInicialCusto
        objRegInventario.dValorEstoque = tSldMesEst.dValorInicial - tSldMesEst.dValorInicialRecIndl
        objRegInventario.dSaldoValorCusto = tSldMesEst.dValorInicialCusto
        
        For iIndice = 1 To iMes - 1
            
            objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tSldMesEst.adQuantEnt(iIndice) - tSldMesEst.adQuantSai(iIndice) - tSldMesEst.adSaldoQuantRecIndl(iIndice)
            objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + tSldMesEst.adValorEnt(iIndice) - tSldMesEst.adValorSai(iIndice) - tSldMesEst.adSaldoValorRecIndl(iIndice)
            
            objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto + tSldMesEst.adSaldoQuantCusto(iIndice)
            objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto + tSldMesEst.adSaldoValorCusto(iIndice)
            
        Next
            
    End If
    
    RegInventario_Obter_Quantidades_SldMesEst = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEst:

    RegInventario_Obter_Quantidades_SldMesEst = gErr

    Select Case gErr
        
        Case 69891, 69892
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69893
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150445)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst1(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim tSldMesEst1 As typeSldMesEst1
Dim tSldMesEst2 As typeSldMesEst2
Dim iMes As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst1

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)
    
    'Lê as quantidade e valores de Material de Terceiro conosco
    sComandoSQL = "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3 "
    'Quantidades e valores
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig3" & CStr(iIndice) & ", " & "SaldoValorConsig3" & CStr(iIndice) & "," & "SaldoQuantDemo3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & "," & "SaldoQuantConserto3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & "," & "SaldoQuantOutros3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & "," & "SaldoQuantBenef3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice)
    Next
    'Tabela, Filtro, Ordem
    sComandoSQL = sComandoSQL & " FROM SldMesEst1 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst1
        lErro = Comando_Executar(alComando(3), sComandoSQL, .dQuantInicialConsig3, .dValorInicialConsig3, .dQuantInicialDemo3, .dValorInicialDemo3, .dQuantInicialConserto3, .dValorInicialConserto3, .dQuantInicialOutros3, .dValorInicialOutros3, .dQuantInicialBenef3, .dValorInicialBenef3, .adSaldoQuantConsig3(1), .adSaldoValorConsig3(1), .adSaldoQuantDemo3(1), .adSaldoValorDemo3(1), .adSaldoQuantConserto3(1), .adSaldoValorConserto3(1), .adSaldoQuantOutros3(1), .adSaldoValorOutros3(1), .adSaldoQuantBenef3(1), .adSaldoValorBenef3(1), .adSaldoQuantConsig3(2), .adSaldoValorConsig3(2), .adSaldoQuantDemo3(2), .adSaldoValorDemo3(2), .adSaldoQuantConserto3(2), .adSaldoValorConserto3(2), .adSaldoQuantOutros3(2), .adSaldoValorOutros3(2), .adSaldoQuantBenef3(2), .adSaldoValorBenef3(2) _
        , .adSaldoQuantConsig3(3), .adSaldoValorConsig3(3), .adSaldoQuantDemo3(3), .adSaldoValorDemo3(3), .adSaldoQuantConserto3(3), .adSaldoValorConserto3(3), .adSaldoQuantOutros3(3), .adSaldoValorOutros3(3), .adSaldoQuantBenef3(3), .adSaldoValorBenef3(3), .adSaldoQuantConsig3(4), .adSaldoValorConsig3(4), .adSaldoQuantDemo3(4), .adSaldoValorDemo3(4), .adSaldoQuantConserto3(4), .adSaldoValorConserto3(4), .adSaldoQuantOutros3(4), .adSaldoValorOutros3(4), .adSaldoQuantBenef3(4), .adSaldoValorBenef3(4), .adSaldoQuantConsig3(5), .adSaldoValorConsig3(5), .adSaldoQuantDemo3(5), .adSaldoValorDemo3(5), .adSaldoQuantConserto3(5), .adSaldoValorConserto3(5), .adSaldoQuantOutros3(5), .adSaldoValorOutros3(5), .adSaldoQuantBenef3(5), .adSaldoValorBenef3(5), .adSaldoQuantConsig3(6), .adSaldoValorConsig3(6), .adSaldoQuantDemo3(6), .adSaldoValorDemo3(6), .adSaldoQuantConserto3(6), .adSaldoValorConserto3(6), .adSaldoQuantOutros3(6), .adSaldoValorOutros3(6), .adSaldoQuantBenef3(6), .adSaldoValorBenef3(6) _
        , .adSaldoQuantConsig3(7), .adSaldoValorConsig3(7), .adSaldoQuantDemo3(7), .adSaldoValorDemo3(7), .adSaldoQuantConserto3(7), .adSaldoValorConserto3(7), .adSaldoQuantOutros3(7), .adSaldoValorOutros3(7), .adSaldoQuantBenef3(7), .adSaldoValorBenef3(7), .adSaldoQuantConsig3(8), .adSaldoValorConsig3(8), .adSaldoQuantDemo3(8), .adSaldoValorDemo3(8), .adSaldoQuantConserto3(8), .adSaldoValorConserto3(8), .adSaldoQuantOutros3(8), .adSaldoValorOutros3(8), .adSaldoQuantBenef3(8), .adSaldoValorBenef3(8), .adSaldoQuantConsig3(9), .adSaldoValorConsig3(9), .adSaldoQuantDemo3(9), .adSaldoValorDemo3(9), .adSaldoQuantConserto3(9), .adSaldoValorConserto3(9), .adSaldoQuantOutros3(9), .adSaldoValorOutros3(9), .adSaldoQuantBenef3(9), .adSaldoValorBenef3(9), .adSaldoQuantConsig3(10), .adSaldoValorConsig3(10), .adSaldoQuantDemo3(10), .adSaldoValorDemo3(10), .adSaldoQuantConserto3(10), .adSaldoValorConserto3(10), .adSaldoQuantOutros3(10), .adSaldoValorOutros3(10), .adSaldoQuantBenef3(10), .adSaldoValorBenef3(10) _
        , .adSaldoQuantConsig3(11), .adSaldoValorConsig3(11), .adSaldoQuantDemo3(11), .adSaldoValorDemo3(11), .adSaldoQuantConserto3(11), .adSaldoValorConserto3(11), .adSaldoQuantOutros3(11), .adSaldoValorOutros3(11), .adSaldoQuantBenef3(11), .adSaldoValorBenef3(11), .adSaldoQuantConsig3(12), .adSaldoValorConsig3(12), .adSaldoQuantDemo3(12), .adSaldoValorDemo3(12), .adSaldoQuantConserto3(12), .adSaldoValorConserto3(12), .adSaldoQuantOutros3(12), .adSaldoValorOutros3(12), .adSaldoQuantBenef3(12), .adSaldoValorBenef3(12), iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69897
    End With

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69898

    'Alteracao Daniel em 23/11/2001
    If lErro = AD_SQL_SEM_DADOS Then
        
        objRegInventario.dQuantBenef3 = 0
        objRegInventario.dQuantConserto3 = 0
        objRegInventario.dQuantConsig3 = 0
        objRegInventario.dQuantDemo3 = 0
        objRegInventario.dQuantOutras3 = 0
    
        objRegInventario.dValorBenef3 = 0
        objRegInventario.dValorConserto3 = 0
        objRegInventario.dValorConsig3 = 0
        objRegInventario.dValorDemo3 = 0
        objRegInventario.dValorOutras3 = 0

    Else
    
        'Quarda as quantidades Iniciais de Terceiros Conosco
        objRegInventario.dQuantBenef3 = tSldMesEst1.dQuantInicialBenef3
        objRegInventario.dQuantConserto3 = tSldMesEst1.dQuantInicialConserto3
        objRegInventario.dQuantConsig3 = tSldMesEst1.dQuantInicialConsig3
        objRegInventario.dQuantDemo3 = tSldMesEst1.dQuantInicialDemo3
        objRegInventario.dQuantOutras3 = tSldMesEst1.dQuantInicialOutros3
    
        objRegInventario.dValorBenef3 = tSldMesEst1.dValorInicialBenef3
        objRegInventario.dValorConserto3 = tSldMesEst1.dValorInicialConserto3
        objRegInventario.dValorConsig3 = tSldMesEst1.dValorInicialConsig3
        objRegInventario.dValorDemo3 = tSldMesEst1.dValorInicialDemo3
        objRegInventario.dValorOutras3 = tSldMesEst1.dValorInicialOutros3
    
        'Acumula as quantidades Iniciais de Terceiros Conosco
        For iIndice = 1 To iMes - 1
    
            objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + tSldMesEst1.adSaldoQuantBenef3(iIndice)
            objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + tSldMesEst1.adSaldoQuantConserto3(iIndice)
            objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + tSldMesEst1.adSaldoQuantConsig3(iIndice)
            objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + tSldMesEst1.adSaldoQuantDemo3(iIndice)
            objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + tSldMesEst1.adSaldoQuantOutros3(iIndice)
    
            objRegInventario.dValorBenef3 = objRegInventario.dValorBenef3 + tSldMesEst1.adSaldoValorBenef3(iIndice)
            objRegInventario.dValorConserto3 = objRegInventario.dValorConserto3 + tSldMesEst1.adSaldoValorConserto3(iIndice)
            objRegInventario.dValorConsig3 = objRegInventario.dValorConsig3 + tSldMesEst1.adSaldoValorConsig3(iIndice)
            objRegInventario.dValorDemo3 = objRegInventario.dValorDemo3 + tSldMesEst1.adSaldoValorDemo3(iIndice)
            objRegInventario.dValorOutras3 = objRegInventario.dValorOutras3 + tSldMesEst1.adSaldoValorOutros3(iIndice)
    
        Next

    End If
    
    RegInventario_Obter_Quantidades_SldMesEst1 = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEst1:

    RegInventario_Obter_Quantidades_SldMesEst1 = gErr

    Select Case gErr
        
        Case 69897, 69898
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69899
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150446)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst2(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim tSldMesEst1 As typeSldMesEst1
Dim tSldMesEst2 As typeSldMesEst2
Dim iMes As Integer
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst2

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)
    
    'Le as Quantidade dos Escaninhos de Nosso em Terceiros
    sComandoSQL = "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef "
    'Quantidades e valores
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig" & CStr(iIndice) & ", " & "SaldoValorConsig" & CStr(iIndice) & "," & "SaldoQuantDemo" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & "," & "SaldoQuantConserto" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & "," & "SaldoQuantOutros" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & "," & "SaldoQuantBenef" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice)
    Next
    'Tabela, Filtro, Ordem
    sComandoSQL = sComandoSQL & " FROM SldMesEst2 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst2
        lErro = Comando_Executar(alComando(2), sComandoSQL, .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, .adSaldoQuantConsig(1), .adSaldoValorConsig(1), .adSaldoQuantDemo(1), .adSaldoValorDemo(1), .adSaldoQuantConserto(1), .adSaldoValorConserto(1), .adSaldoQuantOutros(1), .adSaldoValorOutros(1), .adSaldoQuantBenef(1), .adSaldoValorBenef(1), .adSaldoQuantConsig(2), .adSaldoValorConsig(2), .adSaldoQuantDemo(2), .adSaldoValorDemo(2), .adSaldoQuantConserto(2), .adSaldoValorConserto(2), .adSaldoQuantOutros(2), .adSaldoValorOutros(2), .adSaldoQuantBenef(2), .adSaldoValorBenef(2) _
        , .adSaldoQuantConsig(3), .adSaldoValorConsig(3), .adSaldoQuantDemo(3), .adSaldoValorDemo(3), .adSaldoQuantConserto(3), .adSaldoValorConserto(3), .adSaldoQuantOutros(3), .adSaldoValorOutros(3), .adSaldoQuantBenef(3), .adSaldoValorBenef(3), .adSaldoQuantConsig(4), .adSaldoValorConsig(4), .adSaldoQuantDemo(4), .adSaldoValorDemo(4), .adSaldoQuantConserto(4), .adSaldoValorConserto(4), .adSaldoQuantOutros(4), .adSaldoValorOutros(4), .adSaldoQuantBenef(4), .adSaldoValorBenef(4), .adSaldoQuantConsig(5), .adSaldoValorConsig(5), .adSaldoQuantDemo(5), .adSaldoValorDemo(5), .adSaldoQuantConserto(5), .adSaldoValorConserto(5), .adSaldoQuantOutros(5), .adSaldoValorOutros(5), .adSaldoQuantBenef(5), .adSaldoValorBenef(5), .adSaldoQuantConsig(6), .adSaldoValorConsig(6), .adSaldoQuantDemo(6), .adSaldoValorDemo(6), .adSaldoQuantConserto(6), .adSaldoValorConserto(6), .adSaldoQuantOutros(6), .adSaldoValorOutros(6), .adSaldoQuantBenef(6), .adSaldoValorBenef(6) _
        , .adSaldoQuantConsig(7), .adSaldoValorConsig(7), .adSaldoQuantDemo(7), .adSaldoValorDemo(7), .adSaldoQuantConserto(7), .adSaldoValorConserto(7), .adSaldoQuantOutros(7), .adSaldoValorOutros(7), .adSaldoQuantBenef(7), .adSaldoValorBenef(7), .adSaldoQuantConsig(8), .adSaldoValorConsig(8), .adSaldoQuantDemo(8), .adSaldoValorDemo(8), .adSaldoQuantConserto(8), .adSaldoValorConserto(8), .adSaldoQuantOutros(8), .adSaldoValorOutros(8), .adSaldoQuantBenef(8), .adSaldoValorBenef(8), .adSaldoQuantConsig(9), .adSaldoValorConsig(9), .adSaldoQuantDemo(9), .adSaldoValorDemo(9), .adSaldoQuantConserto(9), .adSaldoValorConserto(9), .adSaldoQuantOutros(9), .adSaldoValorOutros(9), .adSaldoQuantBenef(9), .adSaldoValorBenef(9), .adSaldoQuantConsig(10), .adSaldoValorConsig(10), .adSaldoQuantDemo(10), .adSaldoValorDemo(10), .adSaldoQuantConserto(10), .adSaldoValorConserto(10), .adSaldoQuantOutros(10), .adSaldoValorOutros(10), .adSaldoQuantBenef(10), .adSaldoValorBenef(10) _
        , .adSaldoQuantConsig(11), .adSaldoValorConsig(11), .adSaldoQuantDemo(11), .adSaldoValorDemo(11), .adSaldoQuantConserto(11), .adSaldoValorConserto(11), .adSaldoQuantOutros(11), .adSaldoValorOutros(11), .adSaldoQuantBenef(11), .adSaldoValorBenef(11), .adSaldoQuantConsig(12), .adSaldoValorConsig(12), .adSaldoQuantDemo(12), .adSaldoValorDemo(12), .adSaldoQuantConserto(12), .adSaldoValorConserto(12), .adSaldoQuantOutros(12), .adSaldoValorOutros(12), .adSaldoQuantBenef(12), .adSaldoValorBenef(12), iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69894
    End With

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69895

    'Alteracao Daniel em 23/11/2001
    If lErro = AD_SQL_SEM_DADOS Then

        'Quantidades Iniciais de Nosso em Terceiros
        objRegInventario.dQuantBenef = 0
        objRegInventario.dQuantConserto = 0
        objRegInventario.dQuantConsig = 0
        objRegInventario.dQuantDemo = 0
        objRegInventario.dQuantOutras = 0
    
        objRegInventario.dValorBenef = 0
        objRegInventario.dValorConserto = 0
        objRegInventario.dValorConsig = 0
        objRegInventario.dValorDemo = 0
        objRegInventario.dValorOutras = 0
        
    Else

        'Quantidades Iniciais de Nosso em Terceiros
        objRegInventario.dQuantBenef = tSldMesEst2.dQuantInicialBenef
        objRegInventario.dQuantConserto = tSldMesEst2.dQuantInicialConserto
        objRegInventario.dQuantConsig = tSldMesEst2.dQuantInicialConsig
        objRegInventario.dQuantDemo = tSldMesEst2.dQuantInicialDemo
        objRegInventario.dQuantOutras = tSldMesEst2.dQuantInicialOutros
    
        objRegInventario.dValorBenef = tSldMesEst2.dValorInicialBenef
        objRegInventario.dValorConserto = tSldMesEst2.dValorInicialConserto
        objRegInventario.dValorConsig = tSldMesEst2.dValorInicialConsig
        objRegInventario.dValorDemo = tSldMesEst2.dValorInicialDemo
        objRegInventario.dValorOutras = tSldMesEst2.dValorInicialOutros
    
        'Acumula mês a mês as quantidades Nossa em Terceiros
        For iIndice = 1 To iMes - 1
    
            objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + tSldMesEst2.adSaldoQuantBenef(iIndice)
            objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + tSldMesEst2.adSaldoQuantConserto(iIndice)
            objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + tSldMesEst2.adSaldoQuantConsig(iIndice)
            objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + tSldMesEst2.adSaldoQuantDemo(iIndice)
            objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + tSldMesEst2.adSaldoQuantOutros(iIndice)
    
            objRegInventario.dValorBenef = objRegInventario.dValorBenef + tSldMesEst2.adSaldoValorBenef(iIndice)
            objRegInventario.dValorConserto = objRegInventario.dValorConserto + tSldMesEst2.adSaldoValorConserto(iIndice)
            objRegInventario.dValorConsig = objRegInventario.dValorConsig + tSldMesEst2.adSaldoValorConsig(iIndice)
            objRegInventario.dValorDemo = objRegInventario.dValorDemo + tSldMesEst2.adSaldoValorDemo(iIndice)
            objRegInventario.dValorOutras = objRegInventario.dValorOutras + tSldMesEst2.adSaldoValorOutros(iIndice)
    
        Next
    
    End If

    RegInventario_Obter_Quantidades_SldMesEst2 = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEst2:

    RegInventario_Obter_Quantidades_SldMesEst2 = gErr

    Select Case gErr

        Case 69894, 69895
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69896
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150447)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quant_Mov_UltimoMes(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim iMes As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim tItemMovEstoque As typeItemMovEstoque
Dim tTipoMovEstoque As typeTipoMovEst
Dim tProduto As typeProduto
Dim dFator As Double
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quant_Mov_UltimoMes

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)
    dtDataInicial = "01/" & iMes & "/" & iAno
    dtDataFinal = objRegInventario.dtData
    
    'Inicializa as variaveis
    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tTipoMovEstoque.sEntradaOuSaida = String(STRING_ENTRADAOUSAIDA, 0)
    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Lê o Movimento de 01/mes/ano até a Data do Inventário
    sComandoSQL = "SELECT Produtos.ClasseUM, Produtos.SiglaUMEstoque, MovimentoEstoque.NumIntDoc,  MovimentoEstoque.Produto, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, MovimentoEstoque.Apropriacao, MovimentoEstoque.Data, MovimentoEstoque.Almoxarifado, MovimentoEstoque.Custo, TiposMovimentoEstoque.AtualizaConsig, TiposMovimentoEstoque.AtualizaDemo, TiposMovimentoEstoque.AtualizaConserto, TiposMovimentoEstoque.AtualizaOutras, TiposMovimentoEstoque.AtualizaBenef, TiposMovimentoEstoque.EntradaOuSaida, TiposMovimentoEstoque.Codigo, TiposMovimentoEstoque.ProdutodeTerc, TiposMovimentoEstoque.AtualizaSaldoCusto, TiposMovimentoEstoque.AtualizaNossaDisp, TiposMovimentoEstoque.AtualizaDefeituosa, " & _
                    "TiposMovimentoEstoque.AtualizaIndOutras FROM Produtos, TiposMovimentoEstoque, MovimentoEstoque WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov AND MovimentoEstoque.Produto = Produtos.Codigo AND MovimentoEstoque.FilialEmpresa = ? AND Data >= ? AND Data <= ? AND Produto = ? AND TiposMovimentoEstoque.AtualizaSoLote = 0"
    lErro = Comando_Executar(alComando(4), sComandoSQL, tProduto.iClasseUM, tProduto.sSiglaUMEstoque, tItemMovEstoque.lNumIntDoc, tItemMovEstoque.sProduto, tItemMovEstoque.dQuantidade, tItemMovEstoque.sSiglaUM, tItemMovEstoque.iApropriacao, tItemMovEstoque.dtData, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.dCusto, tTipoMovEstoque.iAtualizaConsig, tTipoMovEstoque.iAtualizaDemo, tTipoMovEstoque.iAtualizaConserto, tTipoMovEstoque.iAtualizaOutras, tTipoMovEstoque.iAtualizaBenef, tTipoMovEstoque.sEntradaOuSaida, tTipoMovEstoque.iCodigo, tTipoMovEstoque.iProdutoDeTerc, tTipoMovEstoque.iAtualizaSaldoCusto, tTipoMovEstoque.iAtualizaNossaDisp, tTipoMovEstoque.iAtualizaDefeituosa, tTipoMovEstoque.iAtualizaIndOutras, objRegInventario.iFilialEmpresa, dtDataInicial, dtDataFinal, objRegInventario.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 69900

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69901

'------------------- Acumula os valores para este Mes ------------------

    Do While lErro = AD_SQL_SUCESSO

        'Fator de conv. de UMs
        lErro = CF("UM_Conversao", tProduto.iClasseUM, tItemMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 69902

        'Verifica em que escaninho ele faz parte
        If tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then

            If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConserto = objRegInventario.dValorConserto + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                objRegInventario.dQuantConserto = objRegInventario.dQuantConserto - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConserto = objRegInventario.dValorConserto - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConsig = objRegInventario.dValorConsig + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                objRegInventario.dQuantConsig = objRegInventario.dQuantConsig - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConsig = objRegInventario.dValorConsig - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorDemo = objRegInventario.dValorDemo + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                objRegInventario.dQuantDemo = objRegInventario.dQuantDemo - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorDemo = objRegInventario.dValorDemo - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorOutras = objRegInventario.dValorOutras + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                objRegInventario.dQuantOutras = objRegInventario.dQuantOutras - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorOutras = objRegInventario.dValorOutras - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorBenef = objRegInventario.dValorBenef + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                objRegInventario.dQuantBenef = objRegInventario.dQuantBenef - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorBenef = objRegInventario.dValorBenef - (tItemMovEstoque.dCusto)
            End If

        ElseIf tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then

            If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConserto3 = objRegInventario.dValorConserto3 + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConserto3 = objRegInventario.dValorConserto3 - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConsig3 = objRegInventario.dValorConsig3 + tItemMovEstoque.dCusto
            ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorConsig3 = objRegInventario.dValorConsig3 - tItemMovEstoque.dCusto
            End If

            If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorDemo3 = objRegInventario.dValorDemo3 + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorDemo3 = objRegInventario.dValorDemo3 - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorOutras3 = objRegInventario.dValorOutras3 + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorOutras3 = objRegInventario.dValorOutras3 - (tItemMovEstoque.dCusto)
            End If

            If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorBenef3 = objRegInventario.dValorBenef3 + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorBenef3 = objRegInventario.dValorBenef3 - (tItemMovEstoque.dCusto)
            End If

        End If

        If tTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            If tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
            End If
        End If
        
        If tTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            If tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto - (tItemMovEstoque.dCusto)
            End If
        End If
        

        If tTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            If tTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto - (tItemMovEstoque.dCusto)
            End If
        End If
        
        If tTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            If tTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + (tItemMovEstoque.dCusto)
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto + tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto + (tItemMovEstoque.dCusto)
            ElseIf tTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dValorEstoque = objRegInventario.dValorEstoque - (tItemMovEstoque.dCusto)
                objRegInventario.dSaldoQuantCusto = objRegInventario.dSaldoQuantCusto - tItemMovEstoque.dQuantidade * dFator
                objRegInventario.dSaldoValorCusto = objRegInventario.dSaldoValorCusto - (tItemMovEstoque.dCusto)
            End If
        End If
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69903

    Loop

    RegInventario_Obter_Quant_Mov_UltimoMes = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quant_Mov_UltimoMes:

    RegInventario_Obter_Quant_Mov_UltimoMes = gErr

    Select Case gErr

        Case 69900, 69901, 69903
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69902
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150448)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Custos(ByVal objRegInventario As ClassRegInventario) As Long
'Devolve o custo do produto e dos Escaninhos

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iAno As Integer
Dim iMes As Integer
Dim objSldMesEst As New ClassSldMesEst
Dim iIndice As Integer
Dim iEncontrado As Integer

On Error GoTo Erro_RegInventario_Obter_Custos

    objProduto.sCodigo = objRegInventario.sProduto
    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)

    'Lê os dados do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 69908

    'De acordo com a apropriação se calcula o Custo no Mês
    If objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Or objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Or objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
        
        objRegInventario.dValorUnitario = 0

        If objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
            
            '??? a divisao deve ser p/dSaldoQuantCusto ? Jones 28/02/02
            If objRegInventario.dSaldoQuantCusto > 0 Then
                objRegInventario.dValorUnitario = objRegInventario.dSaldoValorCusto / objRegInventario.dSaldoQuantCusto
            End If
                
        ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
            
            objSldMesEst.iAno = iAno
            objSldMesEst.iFilialEmpresa = objRegInventario.iFilialEmpresa
            objSldMesEst.sProduto = objRegInventario.sProduto
            
            lErro = CF("SldMesEst_Le_Custos", iMes, objSldMesEst)
            If lErro = SUCESSO Then
                        
                objRegInventario.dValorUnitario = objSldMesEst.dCustoStandard(iMes)
                
            End If
        
        ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
            
            objSldMesEst.iAno = iAno
            objSldMesEst.iFilialEmpresa = objRegInventario.iFilialEmpresa
            objSldMesEst.sProduto = objRegInventario.sProduto
            
            lErro = CF("SldMesEst_Le", objSldMesEst)
            If lErro = SUCESSO Then
            
                iEncontrado = 0
                
                For iIndice = iMes To 1 Step -1
                    If objSldMesEst.dCustoMedio(iIndice) > 0 Then
                        objRegInventario.dValorUnitario = objSldMesEst.dCustoMedio(iIndice)
                        iEncontrado = 1
                        Exit For
                    End If
                Next
                
                If iEncontrado = 0 And objSldMesEst.dQuantInicial > 0 Then
                    objRegInventario.dValorUnitario = objSldMesEst.dValorInicial / objSldMesEst.dQuantInicial
                End If
        
            End If
        
        End If
        
        'Le o Custo para os Escaninhos
        'Conserto
        If objRegInventario.dQuantConserto <> 0 Then
            objRegInventario.dCustoConserto = objRegInventario.dValorConserto / objRegInventario.dQuantConserto
        Else
            objRegInventario.dCustoConserto = 0
        End If

        If objRegInventario.dQuantConserto3 <> 0 Then
            objRegInventario.dCustoConserto3 = objRegInventario.dValorConserto3 / objRegInventario.dQuantConserto3
        Else
            objRegInventario.dCustoConserto3 = 0
        End If

        'Consignação
        If objRegInventario.dQuantConsig <> 0 Then
            objRegInventario.dCustoConsig = objRegInventario.dValorConsig / objRegInventario.dQuantConsig
        Else
            objRegInventario.dCustoConsig = 0
        End If

        If objRegInventario.dQuantConsig3 <> 0 Then
            objRegInventario.dCustoConsig3 = objRegInventario.dValorConsig3 / objRegInventario.dQuantConsig3
        Else
            objRegInventario.dCustoConsig3 = 0
        End If

        'Demostração
        If objRegInventario.dQuantDemo <> 0 Then
            objRegInventario.dCustoDemo = objRegInventario.dValorDemo / objRegInventario.dQuantDemo
        Else
            objRegInventario.dCustoDemo = 0
        End If

        If objRegInventario.dQuantDemo3 <> 0 Then
            objRegInventario.dCustoDemo3 = objRegInventario.dValorDemo3 / objRegInventario.dQuantDemo3
        Else
            objRegInventario.dCustoDemo3 = 0
        End If

        'Outras
        If objRegInventario.dQuantOutras <> 0 Then
            objRegInventario.dCustoOutras = objRegInventario.dValorOutras / objRegInventario.dQuantOutras
        Else
            objRegInventario.dCustoOutras = 0
        End If

        If objRegInventario.dQuantOutras3 <> 0 Then
            objRegInventario.dCustoOutras3 = objRegInventario.dValorOutras3 / objRegInventario.dQuantOutras3
        Else
            objRegInventario.dCustoOutras3 = 0
        End If

        'Beneficiamento
        If objRegInventario.dQuantBenef <> 0 Then
            objRegInventario.dCustoBenef = objRegInventario.dValorBenef / objRegInventario.dQuantBenef
        Else
            objRegInventario.dCustoBenef = 0
        End If

        If objRegInventario.dQuantBenef3 <> 0 Then
            objRegInventario.dCustoBenef3 = objRegInventario.dValorBenef3 / objRegInventario.dQuantBenef3
        Else
            objRegInventario.dCustoBenef3 = 0
        End If

    End If

    RegInventario_Obter_Custos = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Custos:

    RegInventario_Obter_Custos = gErr

    Select Case gErr

        Case 69908, 69909, 69910

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150449)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_Almox(ByVal objRegInventario As ClassRegInventario) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_RegInventario_Obter_Quantidades_Almox

    'Faz a Leitura da Tabela EstoqueProduto a partir do Produto e seu Almoxarifado
    objEstoqueProduto.sProduto = objRegInventario.sProduto
    objEstoqueProduto.iAlmoxarifado = objRegInventario.iAlmoxarifado
    
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 95480
    
    If lErro = 21306 Then gError 95481
        
    'Se a data do EstoqueProduto for maior que a do inventário => ZERA AS QUANTIDADES
    If objEstoqueProduto.dtDataInicial > objRegInventario.dtData Then
        
        'ZERA AS QUANTIDADES...
        With objRegInventario
        
            .dCustoBenef = 0
            .dCustoBenef3 = 0
            .dCustoConserto = 0
            .dCustoConserto3 = 0
            .dCustoConsig = 0
            .dCustoConsig3 = 0
            .dCustoDemo = 0
            .dCustoDemo3 = 0
            .dCustoOutras = 0
            .dCustoOutras3 = 0
            .dQtdeDeTercConosco = 0
            .dQtdeNossaEmTerc = 0
            .dQuantBenef = 0
            .dQuantBenef3 = 0
            .dQuantConserto = 0
            .dQuantConserto3 = 0
            .dQuantConsig = 0
            .dQuantConsig3 = 0
            .dQuantDemo = 0
            .dQuantDemo3 = 0
            .dQuantidadeUMEstoque = 0
            .dQuantOutras = 0
            .dQuantOutras3 = 0
            .dSaldoQuantCusto = 0
            .dSaldoValorCusto = 0
            .dValorBenef = 0
            .dValorBenef3 = 0
            .dValorConserto = 0
            .dValorConserto3 = 0
            .dValorConsig = 0
            .dValorConsig3 = 0
            .dValorDemo = 0
            .dValorDemo3 = 0
            .dValorEstoque = 0
            .dValorOutras = 0
            .dValorOutras3 = 0
            .dValorUnitario = 0
            
        End With
        
    Else
        
        'Le as Quantidades de SldMesEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = RegInventario_Obter_Quantidades_Almox1(objRegInventario)
        If lErro <> SUCESSO Then gError 92968
    
        'Le as Quantidades de SldDiaEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = RegInventario_Obter_Quantidades_Almox2(objRegInventario)
        If lErro <> SUCESSO Then gError 92969
    
    End If

    RegInventario_Obter_Quantidades_Almox = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_Almox:

    RegInventario_Obter_Quantidades_Almox = gErr

    Select Case gErr
        
        Case 92968, 92969, 95480
        
        Case 95481
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150450)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_Almox1(ByVal objRegInventario As ClassRegInventario) As Long
'Le as Quantidades de SldMesEstAlm Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEstAlm As typeSldMesEstAlm
Dim tSldMesEstAlm1 As typeSldMesEstAlm1
Dim tSldMesEstAlm2 As typeSldMesEstAlm2
Dim iMes As Integer
Dim iIndice As Integer
Dim alComando(1 To 3) As Long

On Error GoTo Erro_RegInventario_Obter_Quantidades_Almox1

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69911
    Next

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)

    '----------------------------- Leitura de SldMesEstAlm ------------------
    'Lê as quantidades Inciais, Entrada e Saída para o Produto
    sComandoSQL = "SELECT QuantInicial, ValorInicial "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", " & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice) & ", " & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice)
    Next
    
    sComandoSQL = sComandoSQL & ", QuantInicialRecIndl, ValorInicialRecIndl, SaldoQuantRecIndl1, SaldoQuantRecIndl2, SaldoQuantRecIndl3, SaldoQuantRecIndl4, SaldoQuantRecIndl5, SaldoQuantRecIndl6, SaldoQuantRecIndl7, SaldoQuantRecIndl8, SaldoQuantRecIndl9, SaldoQuantRecIndl10, SaldoQuantRecIndl11, SaldoQuantRecIndl12, " & _
                    "SaldoValorRecIndl1, SaldoValorRecIndl2, SaldoValorRecIndl3, SaldoValorRecIndl4, SaldoValorRecIndl5, SaldoValorRecIndl6, SaldoValorRecIndl7, SaldoValorRecIndl8, SaldoValorRecIndl9, SaldoValorRecIndl10, SaldoValorRecIndl11, SaldoValorRecIndl12 "
    
    'Tabela, Filtro
    sComandoSQL = sComandoSQL & " FROM SldMesEstAlm WHERE Ano = ? AND Almoxarifado = ? AND Produto = ?"

    With tSldMesEstAlm
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantInicial, .dValorInicial, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), _
        .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), _
        .dQuantInicialRecIndl, .dValorInicialRecIndl, .adSaldoQuantRecIndl(1), .adSaldoQuantRecIndl(2), .adSaldoQuantRecIndl(3), .adSaldoQuantRecIndl(4), .adSaldoQuantRecIndl(5), .adSaldoQuantRecIndl(6), .adSaldoQuantRecIndl(7), .adSaldoQuantRecIndl(8), .adSaldoQuantRecIndl(9), .adSaldoQuantRecIndl(10), .adSaldoQuantRecIndl(11), .adSaldoQuantRecIndl(12), _
        .adSaldoValorRecIndl(1), .adSaldoValorRecIndl(2), .adSaldoValorRecIndl(3), .adSaldoValorRecIndl(4), .adSaldoValorRecIndl(5), .adSaldoValorRecIndl(6), .adSaldoValorRecIndl(7), .adSaldoValorRecIndl(8), .adSaldoValorRecIndl(9), .adSaldoValorRecIndl(10), .adSaldoValorRecIndl(11), .adSaldoValorRecIndl(12), _
        iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69912
    End With

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69913

    'Início da Alteração Daniel em 07/01/02
    'Se não encontrou => Busca em EstoqueProduto
    If lErro = AD_SQL_SEM_DADOS Then gError 92967
            
    '------------------------Acumula os valores de SdlMesEstAlm ------------------
    
    objRegInventario.dQuantidadeUMEstoque = tSldMesEstAlm.dQuantInicial - tSldMesEstAlm.dQuantInicialRecIndl
    objRegInventario.dValorEstoque = tSldMesEstAlm.dValorInicial - tSldMesEstAlm.dValorInicialRecIndl

    For iIndice = 1 To iMes - 1
        objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tSldMesEstAlm.adQuantEnt(iIndice) - tSldMesEstAlm.adQuantSai(iIndice) - tSldMesEstAlm.adSaldoQuantRecIndl(iIndice)
        objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + tSldMesEstAlm.adValorEnt(iIndice) - tSldMesEstAlm.adValorSai(iIndice) - tSldMesEstAlm.adSaldoValorRecIndl(iIndice)
    Next
    
        
    '----------------------------- Leitura de SldMesEstAlm2 ------------------

    'Le as Quantidade dos Escaninhos de Nosso em Terceiros
    sComandoSQL = "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef "
    'Quantidades e valores
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig" & CStr(iIndice) & ", " & "SaldoValorConsig" & CStr(iIndice) & "," & "SaldoQuantDemo" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & "," & "SaldoQuantConserto" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & "," & "SaldoQuantOutros" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & "," & "SaldoQuantBenef" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice)
    Next
    'Tabela, Filtro, Ordem
    sComandoSQL = sComandoSQL & " FROM SldMesEstAlm2 WHERE Ano = ? AND Almoxarifado = ? AND Produto = ?"

    With tSldMesEstAlm2
        lErro = Comando_Executar(alComando(2), sComandoSQL, .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, .adSaldoQuantConsig(1), .adSaldoValorConsig(1), .adSaldoQuantDemo(1), .adSaldoValorDemo(1), .adSaldoQuantConserto(1), .adSaldoValorConserto(1), .adSaldoQuantOutros(1), .adSaldoValorOutros(1), .adSaldoQuantBenef(1), .adSaldoValorBenef(1), .adSaldoQuantConsig(2), .adSaldoValorConsig(2), .adSaldoQuantDemo(2), .adSaldoValorDemo(2), .adSaldoQuantConserto(2), .adSaldoValorConserto(2), .adSaldoQuantOutros(2), .adSaldoValorOutros(2), .adSaldoQuantBenef(2), .adSaldoValorBenef(2) _
        , .adSaldoQuantConsig(3), .adSaldoValorConsig(3), .adSaldoQuantDemo(3), .adSaldoValorDemo(3), .adSaldoQuantConserto(3), .adSaldoValorConserto(3), .adSaldoQuantOutros(3), .adSaldoValorOutros(3), .adSaldoQuantBenef(3), .adSaldoValorBenef(3), .adSaldoQuantConsig(4), .adSaldoValorConsig(4), .adSaldoQuantDemo(4), .adSaldoValorDemo(4), .adSaldoQuantConserto(4), .adSaldoValorConserto(4), .adSaldoQuantOutros(4), .adSaldoValorOutros(4), .adSaldoQuantBenef(4), .adSaldoValorBenef(4), .adSaldoQuantConsig(5), .adSaldoValorConsig(5), .adSaldoQuantDemo(5), .adSaldoValorDemo(5), .adSaldoQuantConserto(5), .adSaldoValorConserto(5), .adSaldoQuantOutros(5), .adSaldoValorOutros(5), .adSaldoQuantBenef(5), .adSaldoValorBenef(5), .adSaldoQuantConsig(6), .adSaldoValorConsig(6), .adSaldoQuantDemo(6), .adSaldoValorDemo(6), .adSaldoQuantConserto(6), .adSaldoValorConserto(6), .adSaldoQuantOutros(6), .adSaldoValorOutros(6), .adSaldoQuantBenef(6), .adSaldoValorBenef(6) _
        , .adSaldoQuantConsig(7), .adSaldoValorConsig(7), .adSaldoQuantDemo(7), .adSaldoValorDemo(7), .adSaldoQuantConserto(7), .adSaldoValorConserto(7), .adSaldoQuantOutros(7), .adSaldoValorOutros(7), .adSaldoQuantBenef(7), .adSaldoValorBenef(7), .adSaldoQuantConsig(8), .adSaldoValorConsig(8), .adSaldoQuantDemo(8), .adSaldoValorDemo(8), .adSaldoQuantConserto(8), .adSaldoValorConserto(8), .adSaldoQuantOutros(8), .adSaldoValorOutros(8), .adSaldoQuantBenef(8), .adSaldoValorBenef(8), .adSaldoQuantConsig(9), .adSaldoValorConsig(9), .adSaldoQuantDemo(9), .adSaldoValorDemo(9), .adSaldoQuantConserto(9), .adSaldoValorConserto(9), .adSaldoQuantOutros(9), .adSaldoValorOutros(9), .adSaldoQuantBenef(9), .adSaldoValorBenef(9), .adSaldoQuantConsig(10), .adSaldoValorConsig(10), .adSaldoQuantDemo(10), .adSaldoValorDemo(10), .adSaldoQuantConserto(10), .adSaldoValorConserto(10), .adSaldoQuantOutros(10), .adSaldoValorOutros(10), .adSaldoQuantBenef(10), .adSaldoValorBenef(10) _
        , .adSaldoQuantConsig(11), .adSaldoValorConsig(11), .adSaldoQuantDemo(11), .adSaldoValorDemo(11), .adSaldoQuantConserto(11), .adSaldoValorConserto(11), .adSaldoQuantOutros(11), .adSaldoValorOutros(11), .adSaldoQuantBenef(11), .adSaldoValorBenef(11), .adSaldoQuantConsig(12), .adSaldoValorConsig(12), .adSaldoQuantDemo(12), .adSaldoValorDemo(12), .adSaldoQuantConserto(12), .adSaldoValorConserto(12), .adSaldoQuantOutros(12), .adSaldoValorOutros(12), .adSaldoQuantBenef(12), .adSaldoValorBenef(12), iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69915
    End With

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69916

    If lErro = AD_SQL_SEM_DADOS Then gError 69917
    
    '------------------------Acumula os valores de SdlMesEst2 ------------------
    
    'Quantidades Iniciais de Nosso em Terceiros
    objRegInventario.dQuantBenef = tSldMesEstAlm2.dQuantInicialBenef
    objRegInventario.dQuantConserto = tSldMesEstAlm2.dQuantInicialConserto
    objRegInventario.dQuantConsig = tSldMesEstAlm2.dQuantInicialConsig
    objRegInventario.dQuantDemo = tSldMesEstAlm2.dQuantInicialDemo
    objRegInventario.dQuantOutras = tSldMesEstAlm2.dQuantInicialOutros

    'Acumula mês a mês as quantidades Nossa em Terceiros
    For iIndice = 1 To iMes - 1

        objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + tSldMesEstAlm2.adSaldoQuantBenef(iIndice)
        objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + tSldMesEstAlm2.adSaldoQuantConserto(iIndice)
        objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + tSldMesEstAlm2.adSaldoQuantConsig(iIndice)
        objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + tSldMesEstAlm2.adSaldoQuantDemo(iIndice)
        objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + tSldMesEstAlm2.adSaldoQuantOutros(iIndice)

    Next
    
    '----------------------------- Leitura de SldMesEstAlm1 ------------------
    
    'Lê as quantidade e valores de Material de Terceiro conosco
    sComandoSQL = "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3 "
    'Quantidades e valores
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig3" & CStr(iIndice) & ", " & "SaldoValorConsig3" & CStr(iIndice) & "," & "SaldoQuantDemo3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & "," & "SaldoQuantConserto3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & "," & "SaldoQuantOutros3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & "," & "SaldoQuantBenef3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice)
    Next
    'Tabela, Filtro, Ordem
    sComandoSQL = sComandoSQL & " FROM SldMesEstAlm1 WHERE Ano = ? AND Almoxarifado = ? AND Produto = ?"

    With tSldMesEstAlm1
        lErro = Comando_Executar(alComando(3), sComandoSQL, .dQuantInicialConsig3, .dValorInicialConsig3, .dQuantInicialDemo3, .dValorInicialDemo3, .dQuantInicialConserto3, .dValorInicialConserto3, .dQuantInicialOutros3, .dValorInicialOutros3, .dQuantInicialBenef3, .dValorInicialBenef3, .adSaldoQuantConsig3(1), .adSaldoValorConsig3(1), .adSaldoQuantDemo3(1), .adSaldoValorDemo3(1), .adSaldoQuantConserto3(1), .adSaldoValorConserto3(1), .adSaldoQuantOutros3(1), .adSaldoValorOutros3(1), .adSaldoQuantBenef3(1), .adSaldoValorBenef3(1), .adSaldoQuantConsig3(2), .adSaldoValorConsig3(2), .adSaldoQuantDemo3(2), .adSaldoValorDemo3(2), .adSaldoQuantConserto3(2), .adSaldoValorConserto3(2), .adSaldoQuantOutros3(2), .adSaldoValorOutros3(2), .adSaldoQuantBenef3(2), .adSaldoValorBenef3(2) _
        , .adSaldoQuantConsig3(3), .adSaldoValorConsig3(3), .adSaldoQuantDemo3(3), .adSaldoValorDemo3(3), .adSaldoQuantConserto3(3), .adSaldoValorConserto3(3), .adSaldoQuantOutros3(3), .adSaldoValorOutros3(3), .adSaldoQuantBenef3(3), .adSaldoValorBenef3(3), .adSaldoQuantConsig3(4), .adSaldoValorConsig3(4), .adSaldoQuantDemo3(4), .adSaldoValorDemo3(4), .adSaldoQuantConserto3(4), .adSaldoValorConserto3(4), .adSaldoQuantOutros3(4), .adSaldoValorOutros3(4), .adSaldoQuantBenef3(4), .adSaldoValorBenef3(4), .adSaldoQuantConsig3(5), .adSaldoValorConsig3(5), .adSaldoQuantDemo3(5), .adSaldoValorDemo3(5), .adSaldoQuantConserto3(5), .adSaldoValorConserto3(5), .adSaldoQuantOutros3(5), .adSaldoValorOutros3(5), .adSaldoQuantBenef3(5), .adSaldoValorBenef3(5), .adSaldoQuantConsig3(6), .adSaldoValorConsig3(6), .adSaldoQuantDemo3(6), .adSaldoValorDemo3(6), .adSaldoQuantConserto3(6), .adSaldoValorConserto3(6), .adSaldoQuantOutros3(6), .adSaldoValorOutros3(6), .adSaldoQuantBenef3(6), .adSaldoValorBenef3(6) _
        , .adSaldoQuantConsig3(7), .adSaldoValorConsig3(7), .adSaldoQuantDemo3(7), .adSaldoValorDemo3(7), .adSaldoQuantConserto3(7), .adSaldoValorConserto3(7), .adSaldoQuantOutros3(7), .adSaldoValorOutros3(7), .adSaldoQuantBenef3(7), .adSaldoValorBenef3(7), .adSaldoQuantConsig3(8), .adSaldoValorConsig3(8), .adSaldoQuantDemo3(8), .adSaldoValorDemo3(8), .adSaldoQuantConserto3(8), .adSaldoValorConserto3(8), .adSaldoQuantOutros3(8), .adSaldoValorOutros3(8), .adSaldoQuantBenef3(8), .adSaldoValorBenef3(8), .adSaldoQuantConsig3(9), .adSaldoValorConsig3(9), .adSaldoQuantDemo3(9), .adSaldoValorDemo3(9), .adSaldoQuantConserto3(9), .adSaldoValorConserto3(9), .adSaldoQuantOutros3(9), .adSaldoValorOutros3(9), .adSaldoQuantBenef3(9), .adSaldoValorBenef3(9), .adSaldoQuantConsig3(10), .adSaldoValorConsig3(10), .adSaldoQuantDemo3(10), .adSaldoValorDemo3(10), .adSaldoQuantConserto3(10), .adSaldoValorConserto3(10), .adSaldoQuantOutros3(10), .adSaldoValorOutros3(10), .adSaldoQuantBenef3(10), .adSaldoValorBenef3(10) _
        , .adSaldoQuantConsig3(11), .adSaldoValorConsig3(11), .adSaldoQuantDemo3(11), .adSaldoValorDemo3(11), .adSaldoQuantConserto3(11), .adSaldoValorConserto3(11), .adSaldoQuantOutros3(11), .adSaldoValorOutros3(11), .adSaldoQuantBenef3(11), .adSaldoValorBenef3(11), .adSaldoQuantConsig3(12), .adSaldoValorConsig3(12), .adSaldoQuantDemo3(12), .adSaldoValorDemo3(12), .adSaldoQuantConserto3(12), .adSaldoValorConserto3(12), .adSaldoQuantOutros3(12), .adSaldoValorOutros3(12), .adSaldoQuantBenef3(12), .adSaldoValorBenef3(12), iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69918
    End With

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69919

    If lErro = AD_SQL_SEM_DADOS Then gError 69920
    
    '------------------------Acumula os valores de SdlMesEst1 ------------------
    
    'Quarda as quantidades Iniciais de Terceiros Conosco
    objRegInventario.dQuantBenef3 = tSldMesEstAlm1.dQuantInicialBenef3
    objRegInventario.dQuantConserto3 = tSldMesEstAlm1.dQuantInicialConserto3
    objRegInventario.dQuantConsig3 = tSldMesEstAlm1.dQuantInicialConsig3
    objRegInventario.dQuantDemo3 = tSldMesEstAlm1.dQuantInicialDemo3
    objRegInventario.dQuantOutras3 = tSldMesEstAlm1.dQuantInicialOutros3

    'Acumula as quantidades Iniciais de Terceiros Conosco
    For iIndice = 1 To iMes - 1

        objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + tSldMesEstAlm1.adSaldoQuantBenef3(iIndice)
        objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + tSldMesEstAlm1.adSaldoQuantConserto3(iIndice)
        objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + tSldMesEstAlm1.adSaldoQuantConsig3(iIndice)
        objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + tSldMesEstAlm1.adSaldoQuantDemo3(iIndice)
        objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + tSldMesEstAlm1.adSaldoQuantOutros3(iIndice)

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RegInventario_Obter_Quantidades_Almox1 = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_Almox1:

    RegInventario_Obter_Quantidades_Almox1 = gErr

    Select Case gErr
        
        Case 69911
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69912, 69913, 69915, 69916, 69918, 69919
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM", gErr, iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case 69917, 69920, 92967
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM_INEXISTENTE", gErr, iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150451)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RegInventario_Obter_Quantidades_Almox2(ByVal objRegInventario As ClassRegInventario) As Long
'Le as Quantidades de SldDiaEstAlm Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldDiaEstAlm As typeSldDiaEstAlm
Dim iMes As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_RegInventario_Obter_Quantidades_Almox2

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92970
    Next

    iAno = Year(objRegInventario.dtData)
    iMes = Month(objRegInventario.dtData)
    dtDataInicial = "01/" & iMes & "/" & iAno
    dtDataFinal = objRegInventario.dtData

    '----------------------------- Leitura de SldDiaEstAlm ------------------
    sComandoSQL = "SELECT QuantEntrada, QuantSaida, QuantEntRecIndl, QuantSaiRecIndl, ValorEntrada, ValorSaida, ValorEntRecIndl, ValorSaiRecIndl, " _
    & "QuantEntConsig3 , QuantSaiConsig3, QuantEntDemo3 , QuantSaiDemo3, " _
    & "QuantEntConserto3 , QuantSaiConserto3, QuantEntOutros3 , QuantSaiOutros3, " _
    & "QuantEntBenef3 , QuantSaiBenef3, QuantEntConsig , QuantSaiConsig, " _
    & "QuantEntDemo , QuantSaiDemo, QuantEntConserto , QuantSaiConserto, " _
    & "QuantEntOutros , QuantSaiOutros, QuantEntBenef , QuantSaiBenef "

    'Tabela, Filtro
    sComandoSQL = sComandoSQL & " FROM SldDiaEstAlm WHERE Almoxarifado = ? AND Produto = ? And Data >= ? And Data <= ?"

    With tSldDiaEstAlm
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantEntrada, .dQuantSaida, .dQuantEntRecIndl, .dQuantSaiRecIndl, .dValorEntrada, .dValorSaida, .dValorEntRecIndl, .dValorSaiRecIndl, _
        .dQuantEntConsig3, .dQuantSaiConsig3, .dQuantEntDemo3, .dQuantSaiDemo3, _
        .dQuantEntConserto3, .dQuantSaiConserto3, .dQuantEntOutros3, .dQuantSaiOutros3, _
        .dQuantEntBenef3, .dQuantSaiBenef3, .dQuantEntConsig, .dQuantSaiConsig, _
        .dQuantEntDemo, .dQuantSaiDemo, .dQuantEntConserto, .dQuantSaiConserto, _
        .dQuantEntOutros, .dQuantSaiOutros, .dQuantEntBenef, .dQuantSaiBenef, _
        objRegInventario.iAlmoxarifado, objRegInventario.sProduto, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 92971
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92972

        Do While lErro = AD_SQL_SUCESSO
    
            '------------------------Acumula os valores de SdlDiaEstAlm ------------------
            objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + .dQuantEntrada - .dQuantSaida - (.dQuantEntRecIndl - .dQuantSaiRecIndl)
            objRegInventario.dValorEstoque = objRegInventario.dValorEstoque + .dValorEntrada - .dValorSaida - (.dValorEntRecIndl - .dValorSaiRecIndl)
            
            objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + .dQuantEntBenef - .dQuantSaiBenef
            objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + .dQuantEntConserto - .dQuantSaiConserto
            objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + .dQuantEntConsig - .dQuantSaiConsig
            objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + .dQuantEntDemo - .dQuantSaiDemo
            objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + .dQuantEntOutros - .dQuantSaiOutros
    
            objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + .dQuantEntBenef3 - .dQuantSaiBenef3
            objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + .dQuantEntConserto3 - .dQuantSaiConserto3
            objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + .dQuantEntConsig3 - .dQuantSaiConsig3
            objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + .dQuantEntDemo3 - .dQuantSaiDemo3
            objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + .dQuantEntOutros3 - .dQuantSaiOutros3
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92973
    
        Loop

    End With

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RegInventario_Obter_Quantidades_Almox2 = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_Almox2:

    RegInventario_Obter_Quantidades_Almox2 = gErr

    Select Case gErr
        
        Case 92970
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 92971, 92972, 92973
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150452)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function LivrosFilial_Le_TodosAbertos(ByVal iTributo As Integer, ByVal colLivrosFilial As Collection) As Long
'Lê Todos os Livros Abertos para o Tributo passado

Dim lErro As Long
Dim lComando As Long
Dim tLivrosFilial As typeLivrosFilial
Dim objLivrosFilial As ClassLivrosFilial

On Error GoTo Erro_LivrosFilial_Le_TodosAbertos

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70169

    tLivrosFilial.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Lê os Livros Fiscais com o Tributo passado
    lErro = Comando_Executar(lComando, "SELECT LivrosFilial.CodLivro, LivrosFiscais.Descricao FROM LivrosFilial, LivrosFiscais WHERE LivrosFilial.CodLivro = LivrosFiscais.Codigo AND LivrosFilial.FilialEmpresa = ? AND LivrosFiscais.CodTributo = ?", tLivrosFilial.iCodLivro, tLivrosFilial.sDescricao, giFilialEmpresa, iTributo)
    If lErro <> AD_SQL_SUCESSO Then gError 70170

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70171

    'Se não encontrou nenhum Livro Fiscal, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70172

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objLivrosFilial = New ClassLivrosFilial

        'Passa dados para objLivroFiscal
        objLivrosFilial.iCodLivro = tLivrosFilial.iCodLivro
        objLivrosFilial.sDescricao = tLivrosFilial.sDescricao

        'Adiciona na coleção de Livros Fiscais
        colLivrosFilial.Add objLivrosFilial

        'Busca Próximo Livro Fiscal
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70173

    Loop

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    LivrosFilial_Le_TodosAbertos = SUCESSO

    Exit Function

Erro_LivrosFilial_Le_TodosAbertos:

    LivrosFilial_Le_TodosAbertos = gErr

    Select Case gErr

        Case 70169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70170, 70171, 70173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 70172 'Não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150453)

    End Select

    'Fechamento de comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ApuracaoICMS_Le_IntervaloData(ByVal objApuracao As ClassRegApuracao) As Long
'Procura por ApuraçãoICMS com data Inicial e Final maiores do que as datas passadas

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ApuracaoICMS_Le_IntervaloData
    
    'Abertura de comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70672
    
    'Procura por Apuração ICMS com data inicial e final maiores do que o intervalo de data passado
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM RegApuracaoICMS WHERE DataInicial <= ? AND DataFinal >= ? AND FilialEmpresa =?", lNumIntDoc, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70673
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70674
    
    'Se não encontrou o Registro de Apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70675
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    ApuracaoICMS_Le_IntervaloData = SUCESSO
    
    Exit Function

Erro_ApuracaoICMS_Le_IntervaloData:
    
    ApuracaoICMS_Le_IntervaloData = gErr
    
    Select Case gErr
            
        Case 70672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 70673, 70674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
            
        Case 70675
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150454)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ApuracaoIPI_Le_IntervaloData(ByVal objApuracao As ClassRegApuracao) As Long
'Procura por ApuraçãoIPI com data Inicial e Final maiores do que as datas passadas

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ApuracaoIPI_Le_IntervaloData
    
    'Abertura de comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70714
    
    'Procura por Apuração IPI com data inicial e final maiores do que o intervalo de data passado
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM RegApuracaoIPI WHERE DataInicial <= ? AND DataFinal >= ? AND FilialEmpresa =?", lNumIntDoc, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70715
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70716
    
    'Se não encontrou o Registro de Apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70717
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    ApuracaoIPI_Le_IntervaloData = SUCESSO
    
    Exit Function

Erro_ApuracaoIPI_Le_IntervaloData:
    
    ApuracaoIPI_Le_IntervaloData = gErr
    
    Select Case gErr
            
        Case 70714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 70715, 70716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
            
        Case 70717
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150455)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LivrosFechados_Le_IntervaloData(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Procura por LivroFechado com data Inicial e Final maiores do que as datas passadas

Dim lComando As Long
Dim lErro As Long
Dim iFolha As Integer

On Error GoTo Erro_LivrosFechados_Le_IntervaloData

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70592
    
    'Procura Livro já fechado com intervalo de datas Inicial e Final maior do que as datas passadas
    lErro = Comando_Executar(lComando, "SELECT FolhaInicial FROM LivrosFechados WHERE DataInicial <= ? AND DataFinal >= ? AND FilialEmpresa = ? AND CodLivro = ?", iFolha, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa, objLivroFechado.iCodLivro)
    If lErro <> AD_SQL_SUCESSO Then gError 70593
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70594
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70595
    
    'Guarda número da próxima folha
    objLivroFechado.iFolhaInicial = iFolha
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    LivrosFechados_Le_IntervaloData = SUCESSO
    
    Exit Function
    
Erro_LivrosFechados_Le_IntervaloData:

    LivrosFechados_Le_IntervaloData = gErr
    
    Select Case gErr
    
        Case 70592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 70593, 70594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
            
        Case 70595 'Livro Fechado não encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150456)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function LivrosFilial_Le_IntervaloData(ByVal objLivrosFilial As ClassLivrosFilial) As Long
'Procura por LivrosFilial com data Inicial e Final maiores do que as datas passadas

Dim lComando As Long
Dim lErro As Long
Dim iNumeroProxFolha As Long

On Error GoTo Erro_LivrosFilial_Le_IntervaloData

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70596
    
    'Procura LivrosFilial com intervalo de datas Inicial e Final maior do que as datas passadas
    lErro = Comando_Executar(lComando, "SELECT NumeroProxFolha FROM LivrosFilial WHERE DataInicial <= ? AND DataFinal >= ? AND FilialEmpresa = ? AND CodLivro = ?", iNumeroProxFolha, objLivrosFilial.dtDataInicial, objLivrosFilial.dtDataFinal, objLivrosFilial.iFilialEmpresa, objLivrosFilial.iCodLivro)
    If lErro <> AD_SQL_SUCESSO Then gError 70597
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70598
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70599
    
    'Guarda o número da Folha
    objLivrosFilial.iNumeroProxFolha = iNumeroProxFolha
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    LivrosFilial_Le_IntervaloData = SUCESSO
    
    Exit Function
    
Erro_LivrosFilial_Le_IntervaloData:

    LivrosFilial_Le_IntervaloData = gErr
    
    Select Case gErr
    
        Case 70596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 70597, 70598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
            
        Case 70599 'Livro aberto não encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150457)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function LivrosFiscais_Valida_Data(ByVal dtData As Date) As Long
'Verifica se a data passada como parametro está entre a Data Inicial
'e a Data Final de algum LivroAberto ou Fechado

Dim lErro As Long
Dim lComando As Long
Dim iCodLivro As Integer

On Error GoTo Erro_LivrosFiscais_Valida_Data

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76318
    
    'Verifica se a Data está entre DataInicial e DataFinal de algum registro da tabela LivrosFilial
    lErro = Comando_Executar(lComando, "SELECT CodLivro FROM LivrosFilial WHERE DataInicial<=? AND DataFinal>=?", iCodLivro, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 76319
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76320
    
    'Se não encontrou o LivroFilial
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Verifica se a Data está entre DataInicial e DataFinal de algum registro da tabela LivrosFechados
        lErro = Comando_Executar(lComando, "SELECT CodLivro FROM LivrosFechados WHERE DataInicial<=? AND DataFinal>=?", iCodLivro, dtData, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 76321
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76322
    
        'Se não encontrou ==>erro
        If lErro = AD_SQL_SEM_DADOS Then gError 76323
        
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    LivrosFiscais_Valida_Data = SUCESSO
    
    Exit Function
    
Erro_LivrosFiscais_Valida_Data:

    LivrosFiscais_Valida_Data = gErr
    
    Select Case gErr
        
        Case 76318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76319, 76320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
            
        Case 76321, 76322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
            
        Case 76323
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150458)
            
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RegInventario_Obter_Quantidades_SldMesEstIni(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Iniciais Inclusive dos escaninhos para do ano seguinte a data do Inventario
'??? poderia otimizar p/ler apenas os campos necessarios
Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim tProduto As typeProduto
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEstIni

    iAno = Year(objRegInventario.dtData) + 1
    
'----------------------------- Leitura de SldMesEst ------------------
    'Lê as quantidades Inciais para o Produto
    sComandoSQL = "SELECT QuantInicial, ValorInicial, QuantInicialCusto, ValorInicialCusto, QuantInicialRecIndl, ValorInicialRecIndl FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, _
            .dQuantInicialRecIndl, .dValorInicialRecIndl, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69891
    End With

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69892

    If lErro = AD_SQL_SEM_DADOS Then
        
        objRegInventario.dQuantidadeUMEstoque = 0
        objRegInventario.dSaldoQuantCusto = 0
        objRegInventario.dValorEstoque = 0
        objRegInventario.dSaldoValorCusto = 0
        
    Else
    
        objRegInventario.dQuantidadeUMEstoque = tSldMesEst.dQuantInicial - tSldMesEst.dQuantInicialRecIndl
        objRegInventario.dSaldoQuantCusto = tSldMesEst.dQuantInicialCusto
        objRegInventario.dValorEstoque = tSldMesEst.dValorInicial - tSldMesEst.dValorInicialRecIndl
        objRegInventario.dSaldoValorCusto = tSldMesEst.dValorInicialCusto
    
    End If
        
    RegInventario_Obter_Quantidades_SldMesEstIni = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEstIni:

    RegInventario_Obter_Quantidades_SldMesEstIni = gErr

    Select Case gErr
        
        Case 69891, 69892
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69893
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150459)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst1Ini(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Iniciais Inclusive dos escaninhos para do ano seguinte a data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim tSldMesEst1 As typeSldMesEst1
Dim tSldMesEst2 As typeSldMesEst2
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst1Ini

    iAno = Year(objRegInventario.dtData) + 1
    
    'Lê as quantidade e valores de Material de Terceiro conosco
    sComandoSQL = "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3 "
    'Tabela, Filtro, Ordem
    sComandoSQL = sComandoSQL & " FROM SldMesEst1 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst1
        lErro = Comando_Executar(alComando(3), sComandoSQL, .dQuantInicialConsig3, .dValorInicialConsig3, .dQuantInicialDemo3, .dValorInicialDemo3, .dQuantInicialConserto3, .dValorInicialConserto3, .dQuantInicialOutros3, .dValorInicialOutros3, .dQuantInicialBenef3, .dValorInicialBenef3, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69897
    End With

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69898

    'Alteracao Daniel em 23/11/2001
    If lErro = AD_SQL_SEM_DADOS Then

        objRegInventario.dQuantBenef3 = 0
        objRegInventario.dQuantConserto3 = 0
        objRegInventario.dQuantConsig3 = 0
        objRegInventario.dQuantDemo3 = 0
        objRegInventario.dQuantOutras3 = 0
    
        objRegInventario.dValorBenef3 = 0
        objRegInventario.dValorConserto3 = 0
        objRegInventario.dValorConsig3 = 0
        objRegInventario.dValorDemo3 = 0
        objRegInventario.dValorOutras3 = 0
        
    Else
        
        'Quarda as quantidades Iniciais de Terceiros Conosco
        objRegInventario.dQuantBenef3 = tSldMesEst1.dQuantInicialBenef3
        objRegInventario.dQuantConserto3 = tSldMesEst1.dQuantInicialConserto3
        objRegInventario.dQuantConsig3 = tSldMesEst1.dQuantInicialConsig3
        objRegInventario.dQuantDemo3 = tSldMesEst1.dQuantInicialDemo3
        objRegInventario.dQuantOutras3 = tSldMesEst1.dQuantInicialOutros3
    
        objRegInventario.dValorBenef3 = tSldMesEst1.dValorInicialBenef3
        objRegInventario.dValorConserto3 = tSldMesEst1.dValorInicialConserto3
        objRegInventario.dValorConsig3 = tSldMesEst1.dValorInicialConsig3
        objRegInventario.dValorDemo3 = tSldMesEst1.dValorInicialDemo3
        objRegInventario.dValorOutras3 = tSldMesEst1.dValorInicialOutros3
        
    End If

    RegInventario_Obter_Quantidades_SldMesEst1Ini = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEst1Ini:

    RegInventario_Obter_Quantidades_SldMesEst1Ini = gErr

    Select Case gErr
        
        Case 69897, 69898
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69899
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150460)

    End Select

    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst2Ini(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Iniciais Inclusive dos escaninhos para do ano seguinte a data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEst As typeSldMesEst
Dim tSldMesEst1 As typeSldMesEst1
Dim tSldMesEst2 As typeSldMesEst2
Dim iIndice As Integer

On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst2Ini

    iAno = Year(objRegInventario.dtData) + 1
    
    'Le as Quantidade dos Escaninhos de Nosso em Terceiros
    sComandoSQL = "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef FROM SldMesEst2 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    With tSldMesEst2
        lErro = Comando_Executar(alComando(2), sComandoSQL, .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69894
    End With

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69895

    'Alteracao Daniel em 23/11/2001
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Quantidades Iniciais de Nosso em Terceiros
        objRegInventario.dQuantBenef = 0
        objRegInventario.dQuantConserto = 0
        objRegInventario.dQuantConsig = 0
        objRegInventario.dQuantDemo = 0
        objRegInventario.dQuantOutras = 0
    
        objRegInventario.dValorBenef = 0
        objRegInventario.dValorConserto = 0
        objRegInventario.dValorConsig = 0
        objRegInventario.dValorDemo = 0
        objRegInventario.dValorOutras = 0
    
    Else
    
        'Quantidades Iniciais de Nosso em Terceiros
        objRegInventario.dQuantBenef = tSldMesEst2.dQuantInicialBenef
        objRegInventario.dQuantConserto = tSldMesEst2.dQuantInicialConserto
        objRegInventario.dQuantConsig = tSldMesEst2.dQuantInicialConsig
        objRegInventario.dQuantDemo = tSldMesEst2.dQuantInicialDemo
        objRegInventario.dQuantOutras = tSldMesEst2.dQuantInicialOutros
    
        objRegInventario.dValorBenef = tSldMesEst2.dValorInicialBenef
        objRegInventario.dValorConserto = tSldMesEst2.dValorInicialConserto
        objRegInventario.dValorConsig = tSldMesEst2.dValorInicialConsig
        objRegInventario.dValorDemo = tSldMesEst2.dValorInicialDemo
        objRegInventario.dValorOutras = tSldMesEst2.dValorInicialOutros
        
    End If

    RegInventario_Obter_Quantidades_SldMesEst2Ini = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quantidades_SldMesEst2Ini:

    RegInventario_Obter_Quantidades_SldMesEst2Ini = gErr

    Select Case gErr

        Case 69894, 69895
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case 69896
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, iAno, objRegInventario.iFilialEmpresa, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150461)

    End Select

    Exit Function

End Function

Function FISConfig_Le2(ByVal objFISConfig As ClassFISConfig) As Long

Dim lErro As Long
Dim sConteudo As String
Dim lComando As Long

On Error GoTo Erro_FISConfig_Le2

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 93632

    sConteudo = String(STRING_CONTEUDO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM FisConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, objFISConfig.sCodigo, objFISConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 93633

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93634

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then gError 93635
    
    objFISConfig.sConteudo = sConteudo
    
    'Libera comando
    Call Comando_Fechar(lComando)

    FISConfig_Le2 = SUCESSO

    Exit Function

Erro_FISConfig_Le2:

    FISConfig_Le2 = gErr

    Select Case gErr

        Case 93632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93633, 93634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FISCONFIG", gErr)

        Case 93635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_FIS_CONFIG_NAO_ENCONTRADO", gErr, objFISConfig.sCodigo, objFISConfig.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150462)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LivrosFilialFechados_Le(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Lê dados do Livro Fiscal de código e Filial Empresa passados em objLivroFechado

Dim lErro As Long
Dim alComando(1) As Long
Dim tLivroFechado As typeLivroFechado
Dim iIndice As Integer

On Error GoTo Erro_LivrosFilialFechados_Le

    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 115033
        
    Next
        
    'Lê os dados básicos do Livro Fiscal Fechado com Código, Filial e Número passados em objLivroFechado
    lErro = Comando_Executar(alComando(0), "SELECT DataInicial, DataFinal FROM LivrosFechados WHERE CodLivro = ? AND FilialEmpresa = ? AND NumeroLivro = ?", _
        tLivroFechado.dtDataInicial, tLivroFechado.dtDataFinal, objLivroFechado.iCodLivro, objLivroFechado.iFilialEmpresa, objLivroFechado.iNumeroLivro)
    If lErro <> AD_SQL_SUCESSO Then gError 115036
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 115037

    'Se não encontrou nenhum Livro corrente, busca em LivrosFilial
    If lErro = AD_SQL_SEM_DADOS Then

        'Lê os dados básicos do Livro Fiscal com Código e FilialEmpresa passados em objLivroFechado
        lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", _
            tLivroFechado.dtDataInicial, tLivroFechado.dtDataFinal, objLivroFechado.iCodLivro, objLivroFechado.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 115034
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 115035
        
        If lErro = AD_SQL_SEM_DADOS Then gError 115038
    
    End If

    'Move os dados lidos para objLivroFechado
    objLivroFechado.dtDataInicial = tLivroFechado.dtDataInicial
    objLivroFechado.dtDataFinal = tLivroFechado.dtDataFinal
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        
        Call Comando_Fechar(alComando(iIndice))
        
    Next

    LivrosFilialFechados_Le = SUCESSO

    Exit Function

Erro_LivrosFilialFechados_Le:

    LivrosFilialFechados_Le = gErr

    Select Case gErr

        Case 115033
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 115034, 115035
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 115036, 115037
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
            
        Case 115038
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150463)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
        
    Next
    
    Exit Function

End Function

Function Produto_Obter_Quantidades_Almox(ByVal objRegInventario As ClassRegInventario, ByVal dtDataFim As Date, ByVal dQuantMinima As Double) As Long
'Obtem o estoque na data especificada e se for menor que a quantidade minima (dQuantMinima)
'obtem o estoque na data fim

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim dQuantReservada As Double

On Error GoTo Erro_Produto_Obter_Quantidades_Almox

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 105224
    Next

    'Faz a Leitura da Tabela EstoqueProduto a partir do Produto e seu Almoxarifado
    objEstoqueProduto.sProduto = objRegInventario.sProduto
    objEstoqueProduto.iAlmoxarifado = objRegInventario.iAlmoxarifado
    
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 105163
    
    If lErro = 21306 Then gError 105164
        
    'Soma as reservas nao vinculadas a pedido
    lErro = Comando_Executar(alComando(1), "SELECT SUM(Quantidade)  FROM Reserva WHERE Almoxarifado = ? AND Produto=? AND TipoDoc=0", dQuantReservada, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 105225

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105226

    'Se a data do EstoqueProduto for maior que a do inventário => ZERA AS QUANTIDADES
    If objEstoqueProduto.dtDataInicial > objRegInventario.dtData Then
        
            gError 105168
        
    Else
        
        'Le as Quantidades de SldMesEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = Produto_Obter_Quantidades_Almox1(objRegInventario, objRegInventario.dtData)
        If lErro <> SUCESSO Then gError 105166
    
        'Le as Quantidades de SldDiaEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = Produto_Obter_Quantidades_Almox2(objRegInventario, objRegInventario.dtData)
        If lErro <> SUCESSO Then gError 105167
        
        objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - dQuantReservada
        
        If objRegInventario.dQuantidadeUMEstoque < dQuantMinima Then gError 105165
    
        'Le as Quantidades de SldMesEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = Produto_Obter_Quantidades_Almox1(objRegInventario, dtDataFim)
        If lErro <> SUCESSO Then gError 105169
    
        'Le as Quantidades de SldDiaEstAlm Inclusive dos escaninhos para data do Inventario
        lErro = Produto_Obter_Quantidades_Almox2(objRegInventario, dtDataFim)
        If lErro <> SUCESSO Then gError 105170
    
        objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - dQuantReservada
    
    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Produto_Obter_Quantidades_Almox = SUCESSO

    Exit Function

Erro_Produto_Obter_Quantidades_Almox:

    Produto_Obter_Quantidades_Almox = gErr

    Select Case gErr
        
        Case 105163, 105165, 105166, 105167, 105168, 105169, 105170
        
        Case 105164
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case 105224
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 105225, 105226
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150464)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Produto_Obter_Quantidades_Almox1(ByVal objRegInventario As ClassRegInventario, ByVal dtData As Date) As Long
'Le as Quantidades de SldMesEstAlm Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldMesEstAlm As typeSldMesEstAlm
Dim iMes As Integer
Dim iIndice As Integer
Dim alComando(1 To 3) As Long

On Error GoTo Erro_Produto_Obter_Quantidades_Almox1

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 105174
    Next

    iAno = Year(dtData)
    iMes = Month(dtData)

    '----------------------------- Leitura de SldMesEstAlm ------------------
    'Lê as quantidades Inciais, Entrada e Saída para o Produto
    sComandoSQL = "SELECT QuantInicial "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", " & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice)
    Next
    
    sComandoSQL = sComandoSQL & ", QuantInicialRecIndl, SaldoQuantRecIndl1, SaldoQuantRecIndl2, SaldoQuantRecIndl3, SaldoQuantRecIndl4, SaldoQuantRecIndl5, SaldoQuantRecIndl6, SaldoQuantRecIndl7, SaldoQuantRecIndl8, SaldoQuantRecIndl9, SaldoQuantRecIndl10, SaldoQuantRecIndl11, SaldoQuantRecIndl12 "
                    
    
    'Tabela, Filtro
    sComandoSQL = sComandoSQL & " FROM SldMesEstAlm WHERE Ano = ? AND Almoxarifado = ? AND Produto = ?"

    With tSldMesEstAlm
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantInicial, .adQuantEnt(1), .adQuantSai(1), .adQuantEnt(2), .adQuantSai(2), .adQuantEnt(3), .adQuantSai(3), .adQuantEnt(4), .adQuantSai(4), .adQuantEnt(5), .adQuantSai(5), .adQuantEnt(6), .adQuantSai(6), _
        .adQuantEnt(7), .adQuantSai(7), .adQuantEnt(8), .adQuantSai(8), .adQuantEnt(9), .adQuantSai(9), .adQuantEnt(10), .adQuantSai(10), .adQuantEnt(11), .adQuantSai(11), .adQuantEnt(12), .adQuantSai(12), _
        .dQuantInicialRecIndl, .adSaldoQuantRecIndl(1), .adSaldoQuantRecIndl(2), .adSaldoQuantRecIndl(3), .adSaldoQuantRecIndl(4), .adSaldoQuantRecIndl(5), .adSaldoQuantRecIndl(6), .adSaldoQuantRecIndl(7), .adSaldoQuantRecIndl(8), .adSaldoQuantRecIndl(9), .adSaldoQuantRecIndl(10), .adSaldoQuantRecIndl(11), .adSaldoQuantRecIndl(12), _
        iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 105175
    End With

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105176

    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 105177
            
    '------------------------Acumula os valores de SdlMesEstAlm ------------------
    
    objRegInventario.dQuantidadeUMEstoque = tSldMesEstAlm.dQuantInicial - tSldMesEstAlm.dQuantInicialRecIndl

    For iIndice = 1 To iMes - 1
        objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tSldMesEstAlm.adQuantEnt(iIndice) - tSldMesEstAlm.adQuantSai(iIndice) - tSldMesEstAlm.adSaldoQuantRecIndl(iIndice)
    Next
    

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Produto_Obter_Quantidades_Almox1 = SUCESSO

    Exit Function

Erro_Produto_Obter_Quantidades_Almox1:

    Produto_Obter_Quantidades_Almox1 = gErr

    Select Case gErr
        
        Case 105174
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 105175, 105176
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM", gErr, iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case 105177
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM_INEXISTENTE", gErr, iAno, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150465)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Produto_Obter_Quantidades_Almox2(ByVal objRegInventario As ClassRegInventario, ByVal dtData As Date) As Long
'Le as Quantidades de SldDiaEstAlm Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sComandoSQL As String
Dim iAno As Integer
Dim tSldDiaEstAlm As typeSldDiaEstAlm
Dim iMes As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_Produto_Obter_Quantidades_Almox2

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 105171
    Next

    iAno = Year(dtData)
    iMes = Month(dtData)
    dtDataInicial = "01/" & iMes & "/" & iAno
    dtDataFinal = objRegInventario.dtData

    '----------------------------- Leitura de SldDiaEstAlm ------------------
    sComandoSQL = "SELECT SUM(QuantEntrada), SUM(QuantSaida), SUM(QuantEntRecIndl), SUM(QuantSaiRecIndl) "

    'Tabela, Filtro
    sComandoSQL = sComandoSQL & " FROM SldDiaEstAlm WHERE Almoxarifado = ? AND Produto = ? And Data >= ? And Data <= ?"

    With tSldDiaEstAlm
        lErro = Comando_Executar(alComando(1), sComandoSQL, .dQuantEntrada, .dQuantSaida, .dQuantEntRecIndl, .dQuantSaiRecIndl, _
        objRegInventario.iAlmoxarifado, objRegInventario.sProduto, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 105172
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105173
    
        '------------------------Acumula os valores de SdlDiaEstAlm ------------------
        objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + .dQuantEntrada - .dQuantSaida - (.dQuantEntRecIndl - .dQuantSaiRecIndl)

    End With

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Produto_Obter_Quantidades_Almox2 = SUCESSO

    Exit Function

Erro_Produto_Obter_Quantidades_Almox2:

    Produto_Obter_Quantidades_Almox2 = gErr

    Select Case gErr
        
        Case 105171
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 105172, 105173
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, objRegInventario.iAlmoxarifado, objRegInventario.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150466)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function InventarioTercProd_Le_InventarioTerc(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'carrega objInventarioTerc.colInventarioTercProd a partir da chave em objInventarioTerc

Dim lErro As Long
Dim lComando As Long
Dim tInventarioProd As typeInventarioTercProd
Dim objInventarioTercProd As ClassInventarioTercProd

On Error GoTo Erro_InventarioTercProd_Le_InventarioTerc

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119710

    'inicializa a string p/ receber o cód. do produto
    tInventarioProd.sProduto = String(STRING_PRODUTO, 0)

    'busca os registros da tabela InvnetarioTercProd a partir das chaves passadas como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Produto, QuantTotal FROM InventarioTercProd WHERE FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Data = ? AND CodEscaninho = ? ORDER BY Produto", _
    tInventarioProd.sProduto, tInventarioProd.dQuantTotal, objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
    If lErro <> AD_SQL_SUCESSO Then gError 119711

    'busc o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119712

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119714
    
    'enquanto houver registro
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia o obj
        Set objInventarioTercProd = New ClassInventarioTercProd
    
        'carrega o obj com os dados lidos na TabelaInventarioTercProd
        With tInventarioProd
    
            objInventarioTercProd.iFilialEmpresa = objInventarioTerc.iFilialEmpresa
            objInventarioTercProd.iTipoTerc = objInventarioTerc.iTipoTerc
            objInventarioTercProd.lCodTerc = objInventarioTerc.lCodTerc
            objInventarioTercProd.iFilialTerc = objInventarioTerc.iFilialTerc
            objInventarioTercProd.dtData = objInventarioTerc.dtData
            objInventarioTercProd.iCodEscaninho = objInventarioTerc.iCodEscaninho
            
            objInventarioTercProd.sProduto = .sProduto
            objInventarioTercProd.dQuantTotal = .dQuantTotal
            
        End With
        
        'adiciona o obj na collection
        objInventarioTerc.colInventarioTercProd.Add objInventarioTercProd
    
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119713
    
    Loop
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   InventarioTercProd_Le_InventarioTerc = SUCESSO

    Exit Function

Erro_InventarioTercProd_Le_InventarioTerc:

    InventarioTercProd_Le_InventarioTerc = gErr

    Select Case gErr

        Case 119714  'sem dados

        Case 119710
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 119711, 119712, 119713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERCPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150467)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function InventarioTercProd_Lock_ObjetosVinculados(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'faz locks nas tabelas de Clientes (e filial), Fornecedores (e Filial) e de Produtos
'para manter a integridade de dados entre as tabelas
'TEM QUE SER CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim objInventarioTercProd As ClassInventarioTercProd
Dim objProduto As New ClassProduto
Dim lComando As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim iFilial As Integer

On Error GoTo Erro_InventarioTercProd_Lock_ObjetosVinculados

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 119726

    'verifica qual foi o tipo de terceiro escolhido
    'se foi o cliente
    If objInventarioTerc.iTipoTerc = TIPO_TERC_CLIENTE Then
        
        'faz um lock compartilhado da filial do cliente
        lErro = CF("FilialCliente_Lock", objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc)
        If lErro <> SUCESSO And lErro <> 43050 Then gError 119733
        
        'se não achou, ==> erro
        If lErro = 43050 Then gError 119735
        
    'senão, foi o fornecedor
    Else
                
        'faz um lock compartilhado da filial do fornecedor
        lErro = CF("FilialFornecedor_Lock", objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc)
        If lErro <> SUCESSO And lErro <> 42870 Then gError 119741
        
        'se não encontrou a filial ==> erro
        If lErro = 42870 Then gError 119743
    
    End If

    'p/ cada produto na coleçao
    For Each objInventarioTercProd In objInventarioTerc.colInventarioTercProd
    
        'guarda o cód do produto em objproduto p/ ser lockado
        objProduto.sCodigo = objInventarioTercProd.sProduto
    
        'faz um lock_compartilhado do produto p/ evitar a exclusão do produto
        lErro = CF("Produto_Lock_Shared", lComando, objProduto)
        If lErro <> SUCESSO And lErro <> 55235 Then gError 119745
    
        'Se não encontrou o produto => erro
        If lErro = 55253 Then gError 119747
    
    Next

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    InventarioTercProd_Lock_ObjetosVinculados = SUCESSO
    
    Exit Function

Erro_InventarioTercProd_Lock_ObjetosVinculados:

    InventarioTercProd_Lock_ObjetosVinculados = gErr

    Select Case gErr

        Case 119726
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119745, 119741, 119733 'erros ja tratados

        Case 119735
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objInventarioTerc.iFilialTerc)

        Case 119743
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr)

        Case 119747
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ENCONTRADO", gErr, objInventarioTercProd.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150468)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GuiaICMS_Le(ByVal objGuiasICMS As ClassGuiasICMS) As Long

Dim lErro As Long
Dim lComando As Long
Dim tGuiasICMS As typeGuiasICMS
Dim sTabela As String

On Error GoTo Erro_GuiaICMS_Le

    lComando = Comando_Abrir
    If lComando = 0 Then gError 125233
    
    If objGuiasICMS.iICMSST = MARCADO Then
        sTabela = "GuiasICMSST"
    Else
        sTabela = "GuiasICMS"
    End If
    
    tGuiasICMS.sLocalEntrega = String(STRING_GUIASICMS_LOCALENTREGA, 0)
    tGuiasICMS.sOrgaoArrecadador = String(STRING_GUIASICMS_ORGAOARRECADADOR, 0)
    tGuiasICMS.sCodReceita = String(STRING_GNRICMS_CODRECEITA, 0)
    tGuiasICMS.sCodObrigRecolher = String(STRING_GNRICMS_CODOBRIGRECOLHER, 0)
    
    If sTabela = "GuiasICMSST" Then
        lErro = Comando_Executar(lComando, "SELECT Valor, DataEntrega, OrgaoArrecadador, LocalEntrega, ApuracaoDe, ApuracaoAte, CodReceita, CodObrigRecolher, Vencimento FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ? AND UF = ?", tGuiasICMS.dValor, tGuiasICMS.dtDataEntrega, tGuiasICMS.sOrgaoArrecadador, tGuiasICMS.sLocalEntrega, tGuiasICMS.dtApuracaoDe, tGuiasICMS.dtApuracaoAte, tGuiasICMS.sCodReceita, tGuiasICMS.sCodObrigRecolher, tGuiasICMS.dtDataVencimento, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa, objGuiasICMS.sUF)
    Else
        lErro = Comando_Executar(lComando, "SELECT Valor, DataEntrega, OrgaoArrecadador, LocalEntrega, ApuracaoDe, ApuracaoAte, CodReceita, CodObrigRecolher, Vencimento FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ?", tGuiasICMS.dValor, tGuiasICMS.dtDataEntrega, tGuiasICMS.sOrgaoArrecadador, tGuiasICMS.sLocalEntrega, tGuiasICMS.dtApuracaoDe, tGuiasICMS.dtApuracaoAte, tGuiasICMS.sCodReceita, tGuiasICMS.sCodObrigRecolher, tGuiasICMS.dtDataVencimento, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 125234
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And AD_SQL_SEM_DADOS Then gError 125235
    
    If lErro = AD_SQL_SEM_DADOS Then gError 125236
    
    With tGuiasICMS
    
        objGuiasICMS.dValor = .dValor
        objGuiasICMS.dtDataEntrega = .dtDataEntrega
        objGuiasICMS.sLocalEntrega = .sLocalEntrega
        objGuiasICMS.sOrgaoArrecadador = .sOrgaoArrecadador
        objGuiasICMS.dtApuracaoDe = .dtApuracaoDe
        objGuiasICMS.dtApuracaoAte = .dtApuracaoAte
        objGuiasICMS.sCodReceita = .sCodReceita
        objGuiasICMS.sCodObrigRecolher = .sCodObrigRecolher
        objGuiasICMS.dtVencimento = .dtDataVencimento
        
    End With
        
    Call Comando_Fechar(lComando)
        
    GuiaICMS_Le = SUCESSO
    
    Exit Function
    
Erro_GuiaICMS_Le:

    GuiaICMS_Le = gErr
    
    Select Case gErr
    
        Case 125233
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125234, 125235
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case 125236
            'registro não encontrado no banco
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150469)
            
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Public Function ApuracaoRes_ObtemGuias(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, ByVal colGuias As Collection) As Long
'preenche colGuias com as guias de ICMS ordenadas por data e numero
'ATENCAO: apenas para apoio a adrelvb

Dim lErro As Long, lComando As Long, objGuia As ClassGuiasICMS, tGuia As typeGuiasICMS
'Dim dtDataDePer As Date, dtDataAtePer As Date

On Error GoTo Erro_ApuracaoRes_ObtemGuias

'    dtDataDePer = DateAdd("m", 1, dtDataDe)
'    dtDataAtePer = DateAdd("m", 1, dtDataAte)
    
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 124106
    
    With tGuia
    
        .sLocalEntrega = String(STRING_GUIASICMS_LOCALENTREGA, 0)
        .sNumero = String(STRING_GUIASICMS_NUMERO, 0)
        .sOrgaoArrecadador = String(STRING_GUIASICMS_ORGAOARRECADADOR, 0)
        
    End With
    
    lErro = Comando_Executar(lComando, "SELECT Numero, Data, Valor, OrgaoArrecadador, DataEntrega, LocalEntrega FROM GuiasICMS WHERE FilialEmpresa = ? AND ApuracaoDe >= ? AND ApuracaoAte <= ? ORDER BY Data, Numero", _
        tGuia.sNumero, tGuia.dtData, tGuia.dValor, tGuia.sOrgaoArrecadador, tGuia.dtDataEntrega, tGuia.sLocalEntrega, iFilialEmpresa, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 124107
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124108
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objGuia = New ClassGuiasICMS
        
        With objGuia
        
            .dtData = tGuia.dtData
            .dtDataEntrega = tGuia.dtDataEntrega
            .dValor = tGuia.dValor
            .iFilialEmpresa = iFilialEmpresa
            .sLocalEntrega = tGuia.sLocalEntrega
            .sNumero = tGuia.sNumero
            .sOrgaoArrecadador = tGuia.sOrgaoArrecadador
        
        End With
        
        colGuias.Add objGuia
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124109
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    ApuracaoRes_ObtemGuias = SUCESSO
     
    Exit Function
    
Erro_ApuracaoRes_ObtemGuias:

    ApuracaoRes_ObtemGuias = gErr
     
    Select Case gErr
          
        Case 124106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 124107, 124108, 124109
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIASICMSREL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150470)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function EstoqueNegativo_Verifica(ByVal iOperacao As Integer, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
    If gobjMAT.iESTTestaEstoqueNegativo = MARCADO Then
        EstoqueNegativo_Verifica = EstoqueNegativo_Verifica2(iOperacao, objItemMovEst, objTipoMovEstoque)
    Else
        EstoqueNegativo_Verifica = SUCESSO
    End If
End Function

Function RegInventario_Obter_Quant_EmProcesso(ByVal objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim lErro As Long
Dim sSQL1 As String
Dim sSQL2 As String
Dim tME As typeItemMovEstoque
Dim tProdR As typeProduto
Dim tProdP As typeProduto
Dim dFatorR As Double
Dim dfatorP As Double
Dim iIndice As Integer
Dim colcolRequisicoes As New Collection
Dim colRequisicoes As New Collection
Dim colReqAux As New Collection
Dim colcolProdEnt As New Collection
Dim colProdEnt As New Collection
Dim colProdAux As New Collection
Dim colOps As New Collection
Dim objItemMovReq As ClassItemMovEstoque
Dim objItemMovProd As ClassItemMovEstoque
Dim objItemMovOP As ClassItemMovEstoque
Dim iTipoOrigemAnt As Integer
Dim lNumIntDocOrigemAnt As Long
Dim dQuantEstDev As Double
Dim dQuantEstDevB3 As Double
Dim dQuantEstDevO3 As Double
Dim dtDataAux As Date
Dim dHoraAux As Double
Dim dtDataB3Aux As Date
Dim dHoraB3Aux As Double
Dim dtDataO3Aux As Date
Dim dHoraO3Aux As Double
Dim dQuantReq As Double
Dim dQuantReqB3 As Double
Dim dQuantReqO3 As Double
Dim dValorReq As Double
Dim dValorReqB3 As Double
Dim dValorReqO3 As Double
Dim iAux As Integer
Dim dtDataInicial As Date

On Error GoTo Erro_RegInventario_Obter_Quant_EmProcesso

    dtDataInicial = "01/" & Month(objRegInventario.dtData) & "/" & Year(objRegInventario.dtData)
    
    'Inicializa as variaveis
    tME.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tME.sOPCodigo = String(STRING_OPCODIGO, 0)
    tProdR.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Pega a Classe e UM de estoque para conversão futura
    lErro = Comando_Executar(alComando(7), "SELECT ClasseUM, SiglaUMEstoque FROM Produtos WHERE Codigo = ?", tProdR.iClasseUM, tProdR.sSiglaUMEstoque, objRegInventario.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 200213

    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200214
    
    'Select de requisições, e devoluções e estornos de requisição
    sSQL1 = "SELECT M.Codigo, M.NumIntDoc,  M.Quantidade, M.SiglaUM, M.Data, "
    sSQL1 = sSQL1 & "M.Almoxarifado, M.Custo, M.TipoMov, M.CodigoOP, M.TipoNumIntDocOrigem, M.NumIntDocOrigem, M.Hora FROM Produtos AS P, MovimentoEstoque AS M "
    sSQL1 = sSQL1 & "WHERE M.Produto = P.Codigo AND M.FilialEmpresa = ? AND M.Data <= ? AND M.Produto = ? AND M.TipoNumIntDocOrigem = 5 AND "
    
    sSQL1 = sSQL1 & "M.NumIntDocOrigem IN (SELECT NumIntDoc FROM ItensOPTodos AS I WHERE I.Situacao <> 3 OR "
    sSQL1 = sSQL1 & "(I.Situacao = 3 AND EXISTS (SELECT M1.NumIntDocOrigem FROM MovimentoEstoque AS M1 "
    sSQL1 = sSQL1 & "WHERE M1.TipoMov IN (7,154,252) AND M1.TipoNumIntDocOrigem = 5 AND "
    sSQL1 = sSQL1 & "M1.NumIntDocOrigem = I.NumIntDoc AND M1.Data > ?))) AND "

    sSQL1 = sSQL1 & "M.TipoMov IN (6,19,20,151,152,250,251) ORDER BY M.TipoNumIntDocOrigem DESC, M.NumIntDocOrigem DESC, M.Data DESC, M.Hora DESC"
    
    'Select de produção, e devoluções e estornos de produção
    sSQL2 = "SELECT P.ClasseUM, P.SiglaUMEstoque, M.NumIntDoc,  M.Quantidade, M.SiglaUM, M.Data, "
    sSQL2 = sSQL2 & "M.Almoxarifado, M.Custo, M.TipoMov, M.Hora FROM Produtos AS P, MovimentoEstoque AS M "
    sSQL2 = sSQL2 & "WHERE M.Produto = P.Codigo AND M.FilialEmpresa = ? AND M.Data <= ? AND M.TipoNumIntDocOrigem = ? AND M.NumIntDocOrigem = ? AND "
    sSQL2 = sSQL2 & "M.TipoMov IN (7,15,16,154,155,252,253) ORDER BY M.Data DESC, M.Hora DESC"
    
    'Le as requisições, e devoluções e estornos de requisição
    lErro = Comando_Executar(alComando(5), sSQL1, tME.lCodigo, tME.lNumIntDoc, tME.dQuantidade, tME.sSiglaUM, tME.dtData, tME.iAlmoxarifado, tME.dCusto, tME.iTipoMov, tME.sOPCodigo, tME.iTipoNumIntDocOrigem, tME.lNumIntDocOrigem, tME.dHora, objRegInventario.iFilialEmpresa, objRegInventario.dtData, objRegInventario.sProduto, objRegInventario.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 200215

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200216

    Do While lErro = AD_SQL_SUCESSO
           
        'Fator de conv. de UMs
        lErro = CF("UM_Conversao", tProdR.iClasseUM, tME.sSiglaUM, tProdR.sSiglaUMEstoque, dFatorR)
        If lErro <> SUCESSO Then gError 200217
    
        Set objItemMovReq = New ClassItemMovEstoque
        
        objItemMovReq.lNumIntDoc = tME.lNumIntDoc
        objItemMovReq.dQuantidade = tME.dQuantidade * dFatorR
        objItemMovReq.sSiglaUM = tProdR.sSiglaUMEstoque
        objItemMovReq.dtData = tME.dtData
        objItemMovReq.iAlmoxarifado = tME.iAlmoxarifado
        objItemMovReq.dCusto = tME.dCusto
        objItemMovReq.iTipoMov = tME.iTipoMov
        objItemMovReq.sOPCodigo = tME.sOPCodigo
        objItemMovReq.iTipoNumIntDocOrigem = tME.iTipoNumIntDocOrigem
        objItemMovReq.lNumIntDocOrigem = tME.lNumIntDocOrigem
        objItemMovReq.dtHora = tME.dHora
        objItemMovReq.lCodigo = tME.lCodigo
        
        colRequisicoes.Add objItemMovReq
        
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200218
        
        'Se mudou a origem ou é o último
        If iTipoOrigemAnt <> tME.iTipoNumIntDocOrigem Or lNumIntDocOrigemAnt <> tME.lNumIntDocOrigem Or lErro = AD_SQL_SEM_DADOS Then
            colcolRequisicoes.Add colRequisicoes
            colOps.Add colRequisicoes.Item(1)
            Set colRequisicoes = New Collection
        End If
        
    Loop
    
    'Trata individualmente cada OP
    For Each objItemMovOP In colOps
    
        Set colProdEnt = New Collection

        tME.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tProdP.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

        'Le as produção, e devoluções e estornos de produção
        lErro = Comando_Executar(alComando(6), sSQL2, tProdP.iClasseUM, tProdP.sSiglaUMEstoque, tME.lNumIntDoc, tME.dQuantidade, tME.sSiglaUM, tME.dtData, tME.iAlmoxarifado, tME.dCusto, tME.iTipoMov, tME.dHora, objRegInventario.iFilialEmpresa, objRegInventario.dtData, objItemMovOP.iTipoNumIntDocOrigem, objItemMovOP.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 200219
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200220
    
        Do While lErro = AD_SQL_SUCESSO
            
            'Fator de conv. de UMs
            lErro = CF("UM_Conversao", tProdP.iClasseUM, tME.sSiglaUM, tProdP.sSiglaUMEstoque, dfatorP)
            If lErro <> SUCESSO Then gError 200221
            
            Set objItemMovProd = New ClassItemMovEstoque
            
            objItemMovProd.lNumIntDoc = tME.lNumIntDoc
            objItemMovProd.dQuantidade = tME.dQuantidade * dfatorP
            objItemMovProd.sSiglaUM = tProdP.sSiglaUMEstoque
            objItemMovProd.dtData = tME.dtData
            objItemMovProd.iAlmoxarifado = tME.iAlmoxarifado
            objItemMovProd.dCusto = tME.dCusto
            objItemMovProd.iTipoMov = tME.iTipoMov
            objItemMovProd.dtHora = tME.dHora
            
            colProdEnt.Add objItemMovProd
            
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200222
            
        Loop
        
        colcolProdEnt.Add colProdEnt
    
    Next
        
    'Retira as quantidades estornadas e devolvidas das duas coleções
    iIndice = 0
    For Each colRequisicoes In colcolRequisicoes
        Set colReqAux = New Collection
        iIndice = iIndice + 1
        dQuantEstDev = 0
        dQuantEstDevB3 = 0
        dQuantEstDevO3 = 0
        For Each objItemMovReq In colRequisicoes 'As requisições tem que estar em ordem decrescente de data
        
            Select Case objItemMovReq.iTipoMov
            
                Case 19, 20
                    dQuantEstDev = dQuantEstDev + objItemMovReq.dQuantidade
                Case 152
                    dQuantEstDevB3 = dQuantEstDevB3 + objItemMovReq.dQuantidade
                Case 251
                    dQuantEstDevO3 = dQuantEstDevO3 + objItemMovReq.dQuantidade
                Case 6
                    If dQuantEstDev < objItemMovReq.dQuantidade Then
                        
                        objItemMovReq.dCusto = (objItemMovReq.dCusto / objItemMovReq.dQuantidade) * (objItemMovReq.dQuantidade - dQuantEstDev)
                        objItemMovReq.dQuantidade = objItemMovReq.dQuantidade - dQuantEstDev
                        dQuantEstDev = 0
                        
                        colReqAux.Add objItemMovReq
                        
                    Else
                        dQuantEstDev = dQuantEstDev - objItemMovReq.dQuantidade
                    End If
                
                Case 151
                    If dQuantEstDevB3 < objItemMovReq.dQuantidade Then
                        
                        objItemMovReq.dCusto = (objItemMovReq.dCusto / objItemMovReq.dQuantidade) * (objItemMovReq.dQuantidade - dQuantEstDevB3)
                        objItemMovReq.dQuantidade = objItemMovReq.dQuantidade - dQuantEstDevB3
                        dQuantEstDevB3 = 0
                        
                        colReqAux.Add objItemMovReq
                        
                    Else
                        dQuantEstDevB3 = dQuantEstDevB3 - objItemMovReq.dQuantidade
                    End If
                    
                Case 250
                    If dQuantEstDevO3 < objItemMovReq.dQuantidade Then
                        
                        objItemMovReq.dCusto = (objItemMovReq.dCusto / objItemMovReq.dQuantidade) * (objItemMovReq.dQuantidade - dQuantEstDevO3)
                        objItemMovReq.dQuantidade = objItemMovReq.dQuantidade - dQuantEstDevO3
                        dQuantEstDevO3 = 0
                        
                        colReqAux.Add objItemMovReq
                        
                    Else
                        dQuantEstDevO3 = dQuantEstDevO3 - objItemMovReq.dQuantidade
                    End If
                    
            End Select
    
        Next
        
        For iAux = colcolRequisicoes.Item(iIndice).Count To 1 Step -1
            colcolRequisicoes.Item(iIndice).Remove iAux
        Next
        For iAux = 1 To colReqAux.Count
            colcolRequisicoes.Item(iIndice).Add colReqAux.Item(iAux)
        Next

    Next
        
    iIndice = 0
    For Each colProdEnt In colcolProdEnt
        Set colProdAux = New Collection
        iIndice = iIndice + 1
        dQuantEstDev = 0
        dQuantEstDevB3 = 0
        dQuantEstDevO3 = 0
        For Each objItemMovProd In colProdEnt 'As requisições tem que estar em ordem decrescente de data
        
            Select Case objItemMovProd.iTipoMov
            
                Case 15, 16
                    dQuantEstDev = dQuantEstDev + objItemMovProd.dQuantidade
                Case 155
                    dQuantEstDevB3 = dQuantEstDevB3 + objItemMovProd.dQuantidade
                Case 253
                    dQuantEstDevO3 = dQuantEstDevO3 + objItemMovProd.dQuantidade
                Case 7
                    If dQuantEstDev < objItemMovProd.dQuantidade Then
                        
                        objItemMovProd.dCusto = (objItemMovProd.dCusto / objItemMovProd.dQuantidade) * (objItemMovProd.dQuantidade - dQuantEstDev)
                        objItemMovProd.dQuantidade = objItemMovProd.dQuantidade - dQuantEstDev
                        dQuantEstDev = 0
                        
                        colProdAux.Add objItemMovProd
                        
                    Else
                        dQuantEstDev = dQuantEstDev - objItemMovProd.dQuantidade
                    End If
                
                Case 154
                    If dQuantEstDevB3 < objItemMovProd.dQuantidade Then
                        
                        objItemMovProd.dCusto = (objItemMovProd.dCusto / objItemMovProd.dQuantidade) * (objItemMovProd.dQuantidade - dQuantEstDevB3)
                        objItemMovProd.dQuantidade = objItemMovProd.dQuantidade - dQuantEstDevB3
                        dQuantEstDevB3 = 0
                        
                        colProdAux.Add objItemMovProd
                        
                    Else
                        dQuantEstDevB3 = dQuantEstDevB3 - objItemMovProd.dQuantidade
                    End If
                    
                Case 252
                    If dQuantEstDevO3 < objItemMovProd.dQuantidade Then
                        
                        objItemMovProd.dCusto = (objItemMovProd.dCusto / objItemMovProd.dQuantidade) * (objItemMovProd.dQuantidade - dQuantEstDevO3)
                        objItemMovProd.dQuantidade = objItemMovProd.dQuantidade - dQuantEstDevO3
                        dQuantEstDevO3 = 0
                        
                        colProdAux.Add objItemMovProd
                        
                    Else
                        dQuantEstDevO3 = dQuantEstDevO3 - objItemMovProd.dQuantidade
                    End If
                    
            End Select
    
        Next
        
        For iAux = colcolProdEnt.Item(iIndice).Count To 1 Step -1
            colcolProdEnt.Item(iIndice).Remove iAux
        Next
        For iAux = 1 To colProdAux.Count
            colcolProdEnt.Item(iIndice).Add colProdAux.Item(iAux)
        Next
        
    Next
    
'    'Retira as requisições que já deram entrada
'    iIndice = 0
'    For Each colRequisicoes In colcolRequisicoes
'        iIndice = iIndice + 1
'        dtDataAux = DATA_NULA
'        dtDataB3Aux = DATA_NULA
'        dtDataO3Aux = DATA_NULA
'        Set colReqAux = New Collection
'        Set colProdEnt = colcolProdEnt.Item(iIndice)
'
'        'COLOCAR COMO TRATAR NESSE PONTO NO CONFIG, A PRINCÍPIO EU VOU IGNORAR TODAS REQUISIÇÕES QUE TEM
'        'UMA PRODUÇÃO ENTRADA FEIRA APÓS A REQUISICAO
'
'        For Each objItemMovProd In colProdEnt
'            Select Case objItemMovProd.iTipoMov
'                Case 7
'                    If dtDataAux = DATA_NULA Then
'                        dtDataAux = objItemMovProd.dtData
'                        dHoraAux = objItemMovProd.dtHora
'                    End If
'                Case 154
'                    If dtDataB3Aux = DATA_NULA Then
'                        dtDataB3Aux = objItemMovProd.dtData
'                        dHoraB3Aux = objItemMovProd.dtHora
'                    End If
'                Case 252
'                    If dtDataO3Aux = DATA_NULA Then
'                        dtDataO3Aux = objItemMovProd.dtData
'                        dHoraO3Aux = objItemMovProd.dtHora
'                    End If
'            End Select
'        Next
'
'        For Each objItemMovReq In colRequisicoes
'            Select Case objItemMovReq.iTipoMov
'                Case 6
'                    If objItemMovReq.dtData > dtDataAux Then 'Or (objItemMovReq.dtData = dtDataAux And objItemMovReq.dtHora = dHoraAux) Then
'                        colReqAux.Add objItemMovReq
'                        dQuantReq = dQuantReq + objItemMovReq.dQuantidade
'                        dValorReq = dValorReq + objItemMovReq.dCusto
'                    End If
'                Case 151
'                    If objItemMovReq.dtData > dtDataB3Aux Then 'Or (objItemMovReq.dtData = dtDataB3Aux And objItemMovReq.dtHora = dHoraB3Aux) Then
'                        colReqAux.Add objItemMovReq
'                        dQuantReqB3 = dQuantReqB3 + objItemMovReq.dQuantidade
'                        dValorReqB3 = dValorReqB3 + objItemMovReq.dCusto
'                    End If
'                Case 250
'                    If objItemMovReq.dtData > dtDataO3Aux Then ' Or (objItemMovReq.dtData = dtDataO3Aux And objItemMovReq.dtHora = dHoraO3Aux) Then
'                        colReqAux.Add objItemMovReq
'                        dQuantReqO3 = dQuantReqO3 + objItemMovReq.dQuantidade
'                        dValorReqO3 = dValorReqO3 + objItemMovReq.dCusto
'                    End If
'            End Select
'        Next
'        For iAux = colcolRequisicoes.Item(iIndice).Count To 1 Step -1
'            colcolRequisicoes.Item(iIndice).Remove iAux
'        Next
'        For iAux = 1 To colReqAux.Count
'            colcolRequisicoes.Item(iIndice).Add colReqAux.Item(iAux)
'        Next
'    Next

    'Retira as requisições que já deram entrada
    iIndice = 0
    For Each colRequisicoes In colcolRequisicoes
        iIndice = iIndice + 1
        dtDataAux = DATA_NULA
        dtDataB3Aux = DATA_NULA
        dtDataO3Aux = DATA_NULA
        Set colReqAux = New Collection
        Set colProdEnt = colcolProdEnt.Item(iIndice)
        
        'COLOCAR COMO TRATAR NESSE PONTO NO CONFIG, A PRINCÍPIO EU VOU IGNORAR TODAS REQUISIÇÕES QUE TEM
        'UMA PRODUÇÃO ENTRADA FEIRA APÓS A REQUISICAO
        
        For Each objItemMovProd In colProdEnt
            If dtDataAux = DATA_NULA Then
                dtDataAux = objItemMovProd.dtData
                dHoraAux = objItemMovProd.dtHora
                Exit For
            End If
        Next
        
        For Each objItemMovReq In colRequisicoes
            If objItemMovReq.dtData > dtDataAux Then
            
                colReqAux.Add objItemMovReq
            
                Select Case objItemMovReq.iTipoMov
                    Case 6
                        dQuantReq = dQuantReq + objItemMovReq.dQuantidade
                        dValorReq = dValorReq + objItemMovReq.dCusto
                    Case 151
                        dQuantReqB3 = dQuantReqB3 + objItemMovReq.dQuantidade
                        dValorReqB3 = dValorReqB3 + objItemMovReq.dCusto
                    Case 250
                        dQuantReqO3 = dQuantReqO3 + objItemMovReq.dQuantidade
                        dValorReqO3 = dValorReqO3 + objItemMovReq.dCusto
                End Select
                
            End If
        Next
        For iAux = colcolRequisicoes.Item(iIndice).Count To 1 Step -1
            colcolRequisicoes.Item(iIndice).Remove iAux
        Next
        For iAux = 1 To colReqAux.Count
            colcolRequisicoes.Item(iIndice).Add colReqAux.Item(iAux)
        Next
    Next
    
    objRegInventario.dValorEmProcesso = dValorReq
    objRegInventario.dValorEmProcessoBenef3 = dValorReqB3
    objRegInventario.dValorEmProcessoOutros3 = dValorReqO3
    
    objRegInventario.dQuantEmProcesso = dQuantReq
    objRegInventario.dQuantEmProcessoBenef3 = dQuantReqB3
    objRegInventario.dQuantEmProcessoOutros3 = dValorReqO3
    
    Set objRegInventario.colRequisicoes = New Collection

    For Each colRequisicoes In colcolRequisicoes
        For Each objItemMovReq In colRequisicoes
            objRegInventario.colRequisicoes.Add objItemMovReq
        Next
    Next

    RegInventario_Obter_Quant_EmProcesso = SUCESSO

    Exit Function

Erro_RegInventario_Obter_Quant_EmProcesso:

    RegInventario_Obter_Quant_EmProcesso = gErr

    Select Case gErr
    
        Case 200213, 200214
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 200215, 200216, 200218, 200219, 200220, 200222
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 200217, 200221
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200223)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RegApuracaoPISCofins_Le(ByVal objApurPISCofins As ClassRegApuracaoPISCofins) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApurPISCofins As typeRegApuracaoPISCofins

On Error GoTo Erro_RegApuracaoPISCofins_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213000

    'Alocação de espaço no buffer
    tApurPISCofins.sCNPJCedCred = String(STRING_CGC, 0)
    tApurPISCofins.sCodCred = String(STRING_MAXIMO, 0)

    'Le a tabelaRegApuracaoPISCofins
    If objApurPISCofins.lNumIntDoc <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Tipo, Ano, Mes, OrigCred, CNPJCedCred, " & _
                    "CodCred, VlCredApu, VlCredExtApu, VlTotCredApu, VlCredDescPAAnt, VlCredPerPAAnt, VlCredDCompPAAnt, " & _
                    "SdCredDispEFD, VlCredDescEFD, VlCredPerEFD, VlCredDCompEFD, VlCredTrans, VlCredOut, SdCredFim FROM RegApuracaoPISCofins WHERE NumIntDoc= ? ", _
                    tApurPISCofins.lNumIntDoc, tApurPISCofins.iFilialEmpresa, tApurPISCofins.iTipo, tApurPISCofins.iAno, tApurPISCofins.iMes, _
                    tApurPISCofins.iOrigCred, tApurPISCofins.sCNPJCedCred, tApurPISCofins.sCodCred, tApurPISCofins.dVlCredApu, tApurPISCofins.dVlCredExtApu, _
                    tApurPISCofins.dVlTotCredApu, tApurPISCofins.dVlCredDescPAAnt, tApurPISCofins.dVlCredPerPAAnt, tApurPISCofins.dVlCredDCompPAAnt, tApurPISCofins.dSdCredDispEFD, _
                    tApurPISCofins.dVlCredDescEFD, tApurPISCofins.dVlCredPerEFD, tApurPISCofins.dVlCredDCompEFD, tApurPISCofins.dVlCredTrans, tApurPISCofins.dVlCredOut, _
                    tApurPISCofins.dSdCredFim, objApurPISCofins.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Tipo, Ano, Mes, OrigCred, CNPJCedCred, " & _
                    "CodCred, VlCredApu, VlCredExtApu, VlTotCredApu, VlCredDescPAAnt, VlCredPerPAAnt, VlCredDCompPAAnt, " & _
                    "SdCredDispEFD, VlCredDescEFD, VlCredPerEFD, VlCredDCompEFD, VlCredTrans, VlCredOut, SdCredFim FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ?  AND OrigCred= ?  AND CNPJCedCred= ?  AND CodCred= ? ", _
                    tApurPISCofins.lNumIntDoc, tApurPISCofins.iFilialEmpresa, tApurPISCofins.iTipo, tApurPISCofins.iAno, tApurPISCofins.iMes, _
                    tApurPISCofins.iOrigCred, tApurPISCofins.sCNPJCedCred, tApurPISCofins.sCodCred, tApurPISCofins.dVlCredApu, tApurPISCofins.dVlCredExtApu, _
                    tApurPISCofins.dVlTotCredApu, tApurPISCofins.dVlCredDescPAAnt, tApurPISCofins.dVlCredPerPAAnt, tApurPISCofins.dVlCredDCompPAAnt, tApurPISCofins.dSdCredDispEFD, _
                    tApurPISCofins.dVlCredDescEFD, tApurPISCofins.dVlCredPerEFD, tApurPISCofins.dVlCredDCompEFD, tApurPISCofins.dVlCredTrans, tApurPISCofins.dVlCredOut, _
                    tApurPISCofins.dSdCredFim, objApurPISCofins.iFilialEmpresa, objApurPISCofins.iTipo, objApurPISCofins.iAno, objApurPISCofins.iMes, objApurPISCofins.iOrigCred, objApurPISCofins.sCNPJCedCred, objApurPISCofins.sCodCred)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 213001

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213002

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objApurPISCofins.lNumIntDoc = tApurPISCofins.lNumIntDoc
    objApurPISCofins.iFilialEmpresa = tApurPISCofins.iFilialEmpresa
    objApurPISCofins.iTipo = tApurPISCofins.iTipo
    objApurPISCofins.iAno = tApurPISCofins.iAno
    objApurPISCofins.iMes = tApurPISCofins.iMes
    objApurPISCofins.iOrigCred = tApurPISCofins.iOrigCred
    objApurPISCofins.sCNPJCedCred = tApurPISCofins.sCNPJCedCred
    objApurPISCofins.sCodCred = tApurPISCofins.sCodCred
    objApurPISCofins.dVlCredApu = tApurPISCofins.dVlCredApu
    objApurPISCofins.dVlCredExtApu = tApurPISCofins.dVlCredExtApu
    objApurPISCofins.dVlTotCredApu = tApurPISCofins.dVlTotCredApu
    objApurPISCofins.dVlCredDescPAAnt = tApurPISCofins.dVlCredDescPAAnt
    objApurPISCofins.dVlCredPerPAAnt = tApurPISCofins.dVlCredPerPAAnt
    objApurPISCofins.dVlCredDCompPAAnt = tApurPISCofins.dVlCredDCompPAAnt
    objApurPISCofins.dSdCredDispEFD = tApurPISCofins.dSdCredDispEFD
    objApurPISCofins.dVlCredDescEFD = tApurPISCofins.dVlCredDescEFD
    objApurPISCofins.dVlCredPerEFD = tApurPISCofins.dVlCredPerEFD
    objApurPISCofins.dVlCredDCompEFD = tApurPISCofins.dVlCredDCompEFD
    objApurPISCofins.dVlCredTrans = tApurPISCofins.dVlCredTrans
    objApurPISCofins.dVlCredOut = tApurPISCofins.dVlCredOut
    objApurPISCofins.dSdCredFim = tApurPISCofins.dSdCredFim

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RegApuracaoPISCofins_Le = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Le:

    RegApuracaoPISCofins_Le = gErr

    Select Case gErr

        Case 213000
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213001, 213002
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOPISCOFINS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213003)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function EFDTabelas_Carrega_Combo(ByVal objCombo As Object, ByVal sCodTabela As String, ByVal iCodObrigacao As Integer, ByVal dtData As Date) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sValor As String
Dim sSQL As String, dtDataAtu As Date
Dim sVersao As String
Dim sCodigo As String, sDescricao As String

On Error GoTo Erro_EFDTabelas_Carrega_Combo

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213043
    Next

    'Alocação de espaço no buffer
    sVersao = String(STRING_MAXIMO, 0)
    
    'Obtem a versão da data solicitada
    lErro = Comando_Executar(alComando(1), "SELECT Versao FROM EFDTabelas WHERE CodTabela = ? AND CodObrigacao = ? AND DataAtualizacao < ? ORDER BY DataAtualizacao DESC ", sVersao, sCodTabela, iCodObrigacao, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 213044

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213045
    
    If lErro <> AD_SQL_SUCESSO Then gError 213046
    
    sCodigo = String(STRING_MAXIMO, 0)
    sDescricao = String(STRING_MAXIMO, 0)
    
    'Lê as opções da versão
    lErro = Comando_Executar(alComando(2), "SELECT Codigo, Descricao FROM EFDTabelas WHERE CodTabela = ? AND CodObrigacao = ? AND Versao = ? ORDER BY Codigo ", sCodigo, sDescricao, sCodTabela, iCodObrigacao, sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 213047

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213048
    
    objCombo.Clear

    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem sCodigo & SEPARADOR & sDescricao
  
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213049
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    EFDTabelas_Carrega_Combo = SUCESSO

    Exit Function

Erro_EFDTabelas_Carrega_Combo:

    EFDTabelas_Carrega_Combo = gErr

    Select Case gErr

        Case 213043
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213044, 213045, 213047 To 213049
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EFDTABELAS", gErr)

        Case 213046
            Call Rotina_Erro(vbOKOnly, "ERRO_EFDTABELAS_SEM_DADOS", gErr, sCodTabela, iCodObrigacao, Format(dtData, "dd/mm/yyyy"))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213050)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RegApuracaoPISCofins_Le_SaldoCredorAnt(ByVal iFilialEmpresa As Integer, ByVal iTipo As Integer, ByVal dtData As Date, dSaldoCredor As Double) As Long

Dim lErro As Long
Dim lComando As Long
Dim iAno As Integer, iMes As Integer
Dim dValor As Double, dtDataAux As Date

On Error GoTo Erro_RegApuracaoPISCofins_Le_SaldoCredorAnt

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213058
    
    dSaldoCredor = 0
    
    dtDataAux = DateAdd("m", -1, dtData)
    iAno = Year(dtDataAux)
    iMes = Month(dtDataAux)

    'Le a tabelaRegApuracaoPISCofins
    lErro = Comando_Executar(lComando, "SELECT SUM(SdCredFim) FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ?  ", _
                dValor, iFilialEmpresa, iTipo, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 213059

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213060

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    dSaldoCredor = dValor

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RegApuracaoPISCofins_Le_SaldoCredorAnt = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Le_SaldoCredorAnt:

    RegApuracaoPISCofins_Le_SaldoCredorAnt = gErr

    Select Case gErr

        Case 213058
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213059, 213060
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOPISCOFINS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213061)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueNegativo_Verifica2(ByVal iOperacao As Integer, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
    
Dim bTestar As Boolean
Dim lErro As Long, sComandoSQL As String, tSldDiaEstAlm As typeSldDiaEstAlm
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim testprod As typeEstoqueProduto, objRegInventario As New ClassRegInventario
Dim iIndice As Integer, alComando(1 To 3) As Long
Dim dQuantRecIndl As Double

On Error GoTo Erro_EstoqueNegativo_Verifica2

    bTestar = False
    
    If objTipoMovEstoque.iAtualizaSoLote = 0 Then
    
        'se for um estorno ou uma exclusao ==>
        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or iOperacao = EXCLUSAO_MOVIMENTO Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
        
            If ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP) And objItemMovEst.dQuantidadeEst > 0) Or _
                ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP) And objItemMovEst.dQuantidadeEst < 0) Then bTestar = True
        Else
        
            If ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP) And objItemMovEst.dQuantidadeEst > 0) Or _
                ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP) And objItemMovEst.dQuantidadeEst < 0) Then bTestar = True
                
        End If
    End If
    
    If bTestar Then
    
        'Abre os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 130437
        Next

        'pegar estoque inicial de estoque produto e variacao por slddiaestalm
        With testprod
        
            lErro = Comando_Executar(alComando(1), "SELECT " & _
                "QuantidadeInicial , QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantRecIndl, QuantIndOutras, " & _
                "QuantDefeituosa , QuantConsig3, QuantConsig, QuantDemo3, " & _
                "QuantDemo , QuantConserto3, QuantConserto, QuantOutras3, " & _
                "QuantOutras, QuantBenef, QuantBenef3, " & _
                "QuantInicialConsig3 , QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, " & _
                "QuantInicialConserto3 , QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef, " & _
                "QuantInicialBenef3 FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", _
                .dQuantidadeInicial, .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantRecIndl, .dQuantInd, _
                .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, _
                .dQuantDemo, .dQuantConserto3, .dQuantConserto, .dQuantOutras3, _
                .dQuantOutras, .dQuantBenef, .dQuantBenef3, _
                .dQuantInicialConsig3, .dQuantInicialConsig, .dQuantInicialDemo3, .dQuantInicialDemo, _
                .dQuantInicialConserto3, .dQuantInicialConserto, .dQuantInicialOutras3, .dQuantInicialOutras, .dQuantInicialBenef, _
                .dQuantInicialBenef3, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        End With
            
        If lErro <> AD_SQL_SUCESSO Then gError 130438
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130439
        If lErro <> AD_SQL_SUCESSO Then gError 130440
    
        'atribui estoques iniciais
        With objRegInventario
            .dQuantBenef = testprod.dQuantInicialBenef
            .dQuantBenef3 = testprod.dQuantInicialBenef3
            .dQuantConserto = testprod.dQuantInicialConserto
            .dQuantConserto3 = testprod.dQuantInicialConserto3
            .dQuantConsig = testprod.dQuantInicialConsig
            .dQuantConsig3 = testprod.dQuantInicialConsig3
            .dQuantDemo = testprod.dQuantInicialDemo
            .dQuantDemo3 = testprod.dQuantInicialDemo3
            .dQuantOutras = testprod.dQuantInicialOutras
            .dQuantOutras3 = testprod.dQuantInicialOutras3
            .dQuantidadeUMEstoque = testprod.dQuantidadeInicial + testprod.dQuantInicialBenef3 + testprod.dQuantInicialConserto3 + testprod.dQuantInicialConsig3 + testprod.dQuantInicialDemo3 + testprod.dQuantInicialOutras3
        End With
        
        '----------------------------- Leitura de SldDiaEstAlm ------------------
        sComandoSQL = "SELECT Data, QuantEntrada, QuantSaida, QuantEntRecIndl, QuantSaiRecIndl, ValorEntrada, ValorSaida, ValorEntRecIndl, ValorSaiRecIndl, " _
        & "QuantEntConsig3 , QuantSaiConsig3, QuantEntDemo3 , QuantSaiDemo3, " _
        & "QuantEntConserto3 , QuantSaiConserto3, QuantEntOutros3 , QuantSaiOutros3, " _
        & "QuantEntBenef3 , QuantSaiBenef3, QuantEntConsig , QuantSaiConsig, " _
        & "QuantEntDemo , QuantSaiDemo, QuantEntConserto , QuantSaiConserto, " _
        & "QuantEntOutros , QuantSaiOutros, QuantEntBenef , QuantSaiBenef "
    
        'Tabela, Filtro
        sComandoSQL = sComandoSQL & " FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? ORDER BY Data"
    
        With tSldDiaEstAlm
            lErro = Comando_Executar(alComando(2), sComandoSQL, .dtData, .dQuantEntrada, .dQuantSaida, .dQuantEntRecIndl, .dQuantSaiRecIndl, .dValorEntrada, .dValorSaida, .dValorEntRecIndl, .dValorSaiRecIndl, _
            .dQuantEntConsig3, .dQuantSaiConsig3, .dQuantEntDemo3, .dQuantSaiDemo3, _
            .dQuantEntConserto3, .dQuantSaiConserto3, .dQuantEntOutros3, .dQuantSaiOutros3, _
            .dQuantEntBenef3, .dQuantSaiBenef3, .dQuantEntConsig, .dQuantSaiConsig, _
            .dQuantEntDemo, .dQuantSaiDemo, .dQuantEntConserto, .dQuantSaiConserto, _
            .dQuantEntOutros, .dQuantSaiOutros, .dQuantEntBenef, .dQuantSaiBenef, _
            objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130441
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130442
    
        Do While lErro = AD_SQL_SUCESSO
        
            With tSldDiaEstAlm
                            
                '------------------------Acumula os valores de SdlDiaEstAlm ------------------
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + .dQuantEntrada - .dQuantSaida - (.dQuantEntRecIndl - .dQuantSaiRecIndl)
                
                dQuantRecIndl = dQuantRecIndl + (.dQuantEntRecIndl - .dQuantSaiRecIndl)
                
                objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + .dQuantEntBenef - .dQuantSaiBenef
                objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + .dQuantEntConserto - .dQuantSaiConserto
                objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + .dQuantEntConsig - .dQuantSaiConsig
                objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + .dQuantEntDemo - .dQuantSaiDemo
                objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + .dQuantEntOutros - .dQuantSaiOutros
        
                objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + .dQuantEntBenef3 - .dQuantSaiBenef3
                objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + .dQuantEntConserto3 - .dQuantSaiConserto3
                objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + .dQuantEntConsig3 - .dQuantSaiConsig3
                objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + .dQuantEntDemo3 - .dQuantSaiDemo3
                objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + .dQuantEntOutros3 - .dQuantSaiOutros3
        
            End With
            
            If objRegInventario.dQuantidadeUMEstoque < -QTDE_ESTOQUE_DELTA Then gError 130425
            
            If dQuantRecIndl < -QTDE_ESTOQUE_DELTA Then gError 130426
            
            If objRegInventario.dQuantBenef < -QTDE_ESTOQUE_DELTA Then gError 130427
            If objRegInventario.dQuantConserto < -QTDE_ESTOQUE_DELTA Then gError 130428
            If objRegInventario.dQuantConsig < -QTDE_ESTOQUE_DELTA Then gError 130429
            If objRegInventario.dQuantDemo < -QTDE_ESTOQUE_DELTA Then gError 130430
            If objRegInventario.dQuantOutras < -QTDE_ESTOQUE_DELTA Then gError 130431
            
            If objRegInventario.dQuantBenef3 < -QTDE_ESTOQUE_DELTA Then gError 130432
            If objRegInventario.dQuantConserto3 < -QTDE_ESTOQUE_DELTA Then gError 130433
            If objRegInventario.dQuantConsig3 < -QTDE_ESTOQUE_DELTA Then gError 130434
            If objRegInventario.dQuantDemo3 < -QTDE_ESTOQUE_DELTA Then gError 130435
            If objRegInventario.dQuantOutras3 < -QTDE_ESTOQUE_DELTA Then gError 130436
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130443
    
        Loop
    
        'verificar se os saldos finais acumulados batem com os de estoqueproduto
        '??? If Abs((testprod.dQuantDispNossa + testprod.dQuantReservada) - objRegInventario.dQuantidadeUMEstoque) > QTDE_ESTOQUE_DELTA Then gError 130445
        
        If Abs(testprod.dQuantRecIndl - dQuantRecIndl) > QTDE_ESTOQUE_DELTA Then gError 130446
        
        If Abs(testprod.dQuantBenef - objRegInventario.dQuantBenef) > QTDE_ESTOQUE_DELTA Then gError 130447
        If Abs(testprod.dQuantConserto - objRegInventario.dQuantConserto) > QTDE_ESTOQUE_DELTA Then gError 130448
        If Abs((testprod.dQuantConsig + testprod.dQuantReservadaConsig) - objRegInventario.dQuantConsig) > QTDE_ESTOQUE_DELTA Then gError 130449
        If Abs(testprod.dQuantDemo - objRegInventario.dQuantDemo) > QTDE_ESTOQUE_DELTA Then gError 130450
        If Abs(testprod.dQuantOutras - objRegInventario.dQuantOutras) > QTDE_ESTOQUE_DELTA Then gError 130451
        
        If Abs(testprod.dQuantBenef3 - objRegInventario.dQuantBenef3) > QTDE_ESTOQUE_DELTA Then gError 130452
        If Abs(testprod.dQuantConserto3 - objRegInventario.dQuantConserto3) > QTDE_ESTOQUE_DELTA Then gError 130453
        If Abs(testprod.dQuantConsig3 - objRegInventario.dQuantConsig3) > QTDE_ESTOQUE_DELTA Then gError 130454
        If Abs(testprod.dQuantDemo3 - objRegInventario.dQuantDemo3) > QTDE_ESTOQUE_DELTA Then gError 130455
        If Abs(testprod.dQuantOutras3 - objRegInventario.dQuantOutras3) > QTDE_ESTOQUE_DELTA Then gError 130456
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            lErro = Comando_Fechar(alComando(iIndice))
        Next

    End If
    
    EstoqueNegativo_Verifica2 = SUCESSO
     
    Exit Function
    
Erro_EstoqueNegativo_Verifica2:

    EstoqueNegativo_Verifica2 = gErr
     
    Select Case gErr
          
        Case 130445 To 130456
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_ESTOQUEPRODUTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
            
        Case 130437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130438, 130439, 130440
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        
        Case 130441, 130442
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, objItemMovEst.iAlmoxarifado, objItemMovEst.sProduto)
        
        Case 130427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantBenef)
            
        Case 130428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConserto)
            
        Case 130429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConsig)
            
        Case 130430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantDemo)
            
        Case 130431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantOutras)
            
        Case 130432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF3", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado, objRegInventario.dQuantBenef3)
        
        Case 130433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConserto3)
            
        Case 130434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConsig3)
            
        Case 130435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantDemo3)
            
        Case 130436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantOutras3)
            
        Case 130425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantidadeUMEstoque)
            
        Case 130426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_INDISPONIVEL", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, dQuantRecIndl)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CEST_Le(ByVal objCEST As ClassCEST) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCEST As typeCEST

On Error GoTo Erro_CEST_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 216092

    'Alocação de espaço no buffer
    tCEST.sDescricao = String(STRING_MAXIMO, 0)

    'Le a tabelaCEST
    If Len(Trim(objCEST.sNCM)) <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT LEFT(Descricao,250) FROM CEST WHERE Codigo= ? AND ? LIKE (NCM + '%') ", _
                tCEST.sDescricao, objCEST.sCodigo, objCEST.sNCM)
    Else
        lErro = Comando_Executar(lComando, "SELECT LEFT(Descricao,250) FROM CEST WHERE Codigo= ? AND NCM = ? ", _
                tCEST.sDescricao, objCEST.sCodigo, objCEST.sNCM)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 216093

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216094

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCEST.sDescricao = tCEST.sDescricao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CEST_Le = SUCESSO

    Exit Function

Erro_CEST_Le:

    CEST_Le = gErr

    Select Case gErr

        Case 216092
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 216093, 216094
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CEST", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216095)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function NCM_Le_CEST(ByVal sNCM As String, sCEST As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim sCESTAux As String

On Error GoTo Erro_NCM_Le_CEST

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 216100

    'Alocação de espaço no buffer
    sCESTAux = String(STRING_CEST, 0)

    'Le a tabelaCEST
    lErro = Comando_Executar(lComando, "SELECT MAX(CEST) AS MaxCEST FROM CEST_NCM WHERE NCM = ? GROUP BY NCM HAVING COUNT(*) = 1", _
            sCESTAux, sNCM)
    If lErro <> AD_SQL_SUCESSO Then gError 216101

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216102

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    sCEST = sCESTAux

    'Fecha Comando
    Call Comando_Fechar(lComando)

    NCM_Le_CEST = SUCESSO

    Exit Function

Erro_NCM_Le_CEST:

    NCM_Le_CEST = gErr

    Select Case gErr

        Case 216100
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 216101, 216102
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CEST", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216103)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

