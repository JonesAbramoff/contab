VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Início adicionado por Rafael Menezes 24/09/2002
Public Function ServicosAssociadosComprovantes_Le(ByVal objSolServ As ClassSolicitacaoServico, ByVal objCompServ As ClassCompServ) As Long
'Essa Função verifica se um servico já está associado a um comprovante.
'os passos seguidos são os seguintes:
'checa a tabela de solicitacaoservico para descobrir o lnumintdoc da solicitação de serviço
'checa a tabela de compservgr se já existe um registro com numintdocorigem = lnumintdoc e produto=sproduto

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodigo As Long

On Error GoTo Erro_ServicosAssociadosComprovantes_Le

    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 109100
    
    Next
    
    'busca na tabela SolicitacaoServico, o numintdoc referente a ela
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM SolicitacaoServico WHERE Numero=? AND FilialEmpresa=?", lNumIntDoc, objSolServ.lNumero, objSolServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109101
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109102
    
    'se não encontrar->erro
    If lErro = AD_SQL_SEM_DADOS Then gError 109103
    
    'com o numintdoc encontrado, localizar o código na view SolServServicoAtendidos
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM compservgr WHERE Numintdocorigem=? AND Produto=?", lCodigo, lNumIntDoc, objCompServ.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 109104
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109105
    
    'se não encontrar-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 109106
    
    objCompServ.lCodigo = lCodigo
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next

    ServicosAssociadosComprovantes_Le = SUCESSO
    
    Exit Function

Erro_ServicosAssociadosComprovantes_Le:

    ServicosAssociadosComprovantes_Le = gErr
    
    Select Case gErr
    
        Case 109100
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 109101, 109102
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO", gErr)
            
        Case 109103, 109106
        'serão tratados na rotina chamadora
            
        Case 109104, 109105
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERVGR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next

    Exit Function

End Function
'Fim adicionado por Rafael Menezes 24/09/2002

Public Function CompServGR_Le_CompServItem_Codigo(ByVal objCompServItem As ClassCompServItem) As Long
'Le os CompServItens relacionados com o comprovante em questao
'armazenando-os na colecao

Dim lErro As Long, lComando As Long
Dim sCampos As String, sDoc As String
Dim tCompServItem As typeCompServItem
Dim objDoc As New ClassDocumento

On Error GoTo Erro_CompServGR_Le_CompServItem_Codigo

    'Abrindo Comando..
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99288

    'Inicializando as strings do type
    tCompServItem.sDocExtNumero = String(STRING_DOC_NUMERO, 0)
    tCompServItem.sDocIntNumero = String(STRING_DOC_NUMERO, 0)
    tCompServItem.sMotorista = String(STRING_MOTORISTA, 0)
    tCompServItem.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServItem.sPlacaCaminhao = String(STRING_PLACA, 0)
    tCompServItem.sPlacaCarreta = String(STRING_PLACA, 0)
         
    'Para auxiliar o select... guardo os campos em uma string
    sCampos = "NumIntDoc,NumIntDocOrigem,CodItemServico,DataPrev, " _
    & "HoraPrev,DataInicio,HoraInicio,DataFim,HoraFim,DocIntTipo,DocIntNumero, " _
    & "DocIntDataEmissao,DocExtTipo,DocExtNumero,DocExtDataEmissao,DocExtDataRec, " _
    & "DocExtHoraRec,PlacaCaminhao,PlacaCarreta,Motorista,Observacao"
    
    With tCompServItem
    
        'obtem um conjunto de registros da tabela compservitemgr
        'relacionados com o compserv
        lErro = Comando_Executar(lComando, "SELECT " & sCampos & " FROM CompServItemGR WHERE CodItemServico = ? AND NumIntDocOrigem=?", _
        .lNumIntDoc, .lNumIntDocOrigem, .iCodItemServico, .dtDataPrev, .dHoraPrev, .dtDataInicio, .dHoraInicio, .dtDataFim, .dHoraFim, _
        .iDocIntTipo, .sDocIntNumero, .dtDocIntDataEmissao, .iDocExtTipo, .sDocExtNumero, .dtDocExtDataEmissao, .dtDocExtDataRec, .dDocExtHoraRec, .sPlacaCaminhao, .sPlacaCarreta, .sMotorista, .sObservacao, objCompServItem.iCodItemServico, objCompServItem.lNumIntDocOrigem)
       
    End With
        
    If lErro <> SUCESSO Then gError 99289
        
    'posiciona-se no primeiro registro do conjunto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99290
    
    With objCompServItem
        
        'carrega os dados no obj instanciado anteriormente..
        .dtDataFim = tCompServItem.dtDataFim
        .dtDataInicio = tCompServItem.dtDataInicio
        .dtDataPrev = tCompServItem.dtDataPrev
        .dtDocExtDataEmissao = tCompServItem.dtDocExtDataEmissao
        .dtDocExtDataRec = tCompServItem.dtDocExtDataRec
        .dtDocExtHoraRec = tCompServItem.dDocExtHoraRec
        .dtDocIntDataEmissao = tCompServItem.dtDocIntDataEmissao
        .dtHoraFim = tCompServItem.dHoraFim
        .dtHoraInicio = tCompServItem.dHoraInicio
        .dtHoraPrev = tCompServItem.dHoraPrev
        .iCodItemServico = tCompServItem.iCodItemServico
        .lNumIntDoc = tCompServItem.lNumIntDoc
        .lNumIntDocOrigem = tCompServItem.lNumIntDocOrigem
        .iDocExtTipo = tCompServItem.iDocExtTipo
        .iDocIntTipo = tCompServItem.iDocIntTipo
        .sDocExtNumero = tCompServItem.sDocExtNumero
        .sDocIntNumero = tCompServItem.sDocIntNumero
        
            If tCompServItem.iDocExtTipo <> 0 Then
        
            objDoc.iCodigo = tCompServItem.iDocExtTipo
                        
            'le o documento externo para obter o nomereduzido
            lErro = CF("Documento_Le", objDoc)
            If lErro <> SUCESSO Then gError 99291
                        
            .sDocExtTipo = objDoc.sNomeReduzido
                              
        End If
        
        If tCompServItem.iDocIntTipo <> 0 Then
        
            objDoc.iCodigo = tCompServItem.iDocIntTipo
            
            'le o documento interno para obter o nomereduzido
            lErro = CF("Documento_Le", objDoc)
            If lErro <> SUCESSO Then gError 99292
                        
            .sDocIntTipo = objDoc.sNomeReduzido
        
        End If
        
        .sMotorista = tCompServItem.sMotorista
        .sObservacao = tCompServItem.sObservacao
        .sPlacaCaminhao = tCompServItem.sPlacaCaminhao
        .sPlacaCarreta = tCompServItem.sPlacaCarreta
          
    End With
        
    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServGR_Le_CompServItem_Codigo = SUCESSO

    Exit Function

Erro_CompServGR_Le_CompServItem_Codigo:

    CompServGR_Le_CompServItem_Codigo = gErr

    Select Case gErr
        
        Case 99288
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 99289, 99290
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERVITEM", gErr)
        
        Case 99291, 99292
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ServicoItemServico_Le(objServico As ClassServico) As Long
'le todos as relacoes de Servicos X Itens de Servico que estao no bd

Dim lComando As Long
Dim tServItemServ As typeServItemServ
Dim lErro As Long
Dim objServItemServ As ClassServItemServ

On Error GoTo Erro_ServicoItemServico_Le

    'abre conexao com o bd
    lComando = Comando_Abrir
    If lComando = 0 Then gError 97540
    
    lErro = Comando_Executar(lComando, "SELECT CodItemServico, Ordem FROM ServicoItemServico WHERE Produto = ? Order by Ordem, CodItemServico", tServItemServ.iCodItemServico, tServItemServ.iOrdem, objServico.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 97541

    'busca o primeiro item da relacao resultado da operacao anterior...
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97542

    If lErro = AD_SQL_SEM_DADOS Then gError 97543

    'vai lendo os itens ate que leia todos os selecionados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objServItemServ = New ClassServItemServ
        
        objServItemServ.sProduto = objServico.sProduto
        objServItemServ.iCodItemServico = tServItemServ.iCodItemServico
        objServItemServ.iOrdem = tServItemServ.iOrdem
        
        'adiciona na colecao o obj...
        objServico.colServItemServ.Add objServItemServ

        'busca o proximo selecionado...
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97544

    Loop

    Call Comando_Fechar(lComando)

    ServicoItemServico_Le = SUCESSO

    Exit Function

Erro_ServicoItemServico_Le:

    ServicoItemServico_Le = gErr

    Select Case gErr

        Case 97540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97541, 97542, 97544
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERVICOITEMSERVICO", gErr)

        Case 97543

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemServico_Lock(objItemServico As ClassItemServico) As Long
'Lê e faz o "lock" no item de servico passado por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim sDocumento As String

On Error GoTo Erro_ItemServico_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97518

    'Inicializa as Variáveis do tipo String
    sDescricao = String(STRING_ITEMSERVICO_DESCRICAO, 0)
    
    'Lê o Item de Servico
    lErro = Comando_ExecutarLockado(lComando, "SELECT  Descricao FROM ItemServico WHERE Codigo = ?", sDescricao, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97519

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97520

    'Se não encontrou -> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97521

    'Faz o "lock"
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 97522

    'Guarda so dados lidos em objItemServico
    objItemServico.sDescricao = sDescricao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemServico_Lock = SUCESSO

    Exit Function

Erro_ItemServico_Lock:

    ItemServico_Lock = gErr

    Select Case gErr

        Case 97518
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97519, 97520
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO", gErr, objItemServico.iCodigo)

        Case 97521
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_INEXISTENTE", gErr, objItemServico.iCodigo)

        Case 97522
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMSERVICO", gErr, objItemServico.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemServico_Le(objItemServico As ClassItemServico) As Long
'Le os dados do item, cujo código foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim sDocumento As String
Dim iTipoData As Integer

On Error GoTo Erro_ItemServico_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97031

    sDescricao = String(STRING_ITEMSERVICO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Descricao,DataInicioFim FROM ItemServico WHERE Codigo = ? ", sDescricao, iTipoData, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97032
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97033

    'Item nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 97035

    'carrega os dados no obj
    objItemServico.sDescricao = sDescricao
    objItemServico.iTipoData = iTipoData
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemServico_Le = SUCESSO

    Exit Function

Erro_ItemServico_Le:

    ItemServico_Le = gErr

    Select Case gErr

        Case 97031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97032, 97033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO", gErr, objItemServico.iCodigo)

        Case 97035 'nao achou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemServico_Le_Todos(colItens As Collection) As Long
'le todos os Itens de Servico que estao no bd

Dim lComando As Long
Dim objItemServico As ClassItemServico
Dim iCod As Integer
Dim sDesc As String
Dim lErro As Long

On Error GoTo Erro_ItemServico_Le_Todos

    'abre conexao com o bd
    lComando = Comando_Abrir
    If lComando = 0 Then gError 97003

    'seleciona os itens de servico
    sDesc = String(STRING_ITEMSERVICO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT codigo, descricao FROM ItemServico order by Codigo", iCod, sDesc)
    If lErro <> AD_SQL_SUCESSO Then gError 97004

    'busca o primeiro item da relacao resultado da operacao anterior...
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97005
    
    If lErro = AD_SQL_SEM_DADOS Then gError 97007
    
    'vai lendo os itens ate que leia todos os selecionados
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'carrega no obj os dados do item
        Set objItemServico = New ClassItemServico
        
        objItemServico.iCodigo = iCod
        objItemServico.sDescricao = sDesc
        
        'adiciona na colecao o obj...
        colItens.Add objItemServico
        
        'busca o proximo selecionado...
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97006
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    ItemServico_Le_Todos = SUCESSO
    
    Exit Function

Erro_ItemServico_Le_Todos:

    ItemServico_Le_Todos = gErr
    
    Select Case gErr
    
        Case 97003
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case 97004 To 97006
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO_TODOS", gErr)
        
        Case 97007
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoEmbalagem_Le_Todos(colTipoEmbalagem As Collection) As Long
'Lê todos os Tipos de carga existentes no BD

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim iTipo As Integer
Dim objTipoEmbalagem As ClassTipoEmbalagem

On Error GoTo Erro_TipoEmbalagem_Le_Todos

    lComando = Comando_Abrir
    If lComando = 0 Then gError 96549
    
    sDescricao = String(STRING_TIPOEMBALAGEM_DESCRICAO, 0)
    
    'Vai Buscar no BD os Tipos de Cargas existentes
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao FROM TipoEmbalagem ORDER BY Tipo", iTipo, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 96502
    
    ' Busca o primeiro Tipo de Carga existente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96503
                
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTipoEmbalagem = New ClassTipoEmbalagem
        
        objTipoEmbalagem.iTipo = iTipo
        objTipoEmbalagem.sDescricao = sDescricao
        
        'Coloca o Tipo de Carga na Coleção colTipoEmbalagem
        colTipoEmbalagem.Add objTipoEmbalagem
        
        'Busca o próximo Tipo de Carga no BD
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96504
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    TipoEmbalagem_Le_Todos = SUCESSO
        
    Exit Function
    
Erro_TipoEmbalagem_Le_Todos:

    TipoEmbalagem_Le_Todos = gErr
    
    Select Case gErr

        Case 96502, 96503, 96504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM1", gErr)
            
        Case 96549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoEmbalagem_Le(objTipoEmbalagem As ClassTipoEmbalagem) As Long
'Lê o Tipo de Carga relacionado ao código passado no objTipoEmbalagem

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String

On Error GoTo Erro_TipoEmbalagem_Le

    lComando = Comando_Abrir
    If lComando = 0 Then gError 96550
             
    sDescricao = String(STRING_TIPOEMBALAGEM_DESCRICAO, 0)
    
    'Verifica se existe no BD O Tipo de Carga com o Código passado
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM TipoEmbalagem WHERE Tipo = ?", sDescricao, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96505
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96506
    
    'Se o Código não existe no BD, Erro.
    If lErro = AD_SQL_SEM_DADOS Then gError 96507
           
    objTipoEmbalagem.sDescricao = sDescricao
                  
    Call Comando_Fechar(lComando)
        
    TipoEmbalagem_Le = SUCESSO
        
    Exit Function
    
Erro_TipoEmbalagem_Le:

    TipoEmbalagem_Le = gErr
    
    Select Case gErr

        Case 96505, 96506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
        
        Case 96507
        
        Case 96550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoContainer_Le(objTipoContainer As ClassTipoContainer) As Long
'Le os dados do tipo, cujo código foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim dValor As Double
Dim dIso As Double

On Error GoTo Erro_TipoContainer_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97083

    sDescricao = String(STRING_TIPOCONTAINER_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Descricao,Valor,ISO FROM TipoContainer WHERE Tipo = ? ", sDescricao, dValor, dIso, objTipoContainer.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 97084
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97085

    'Item nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 97086

    'carrega os dados no obj
    objTipoContainer.dValor = dValor
    objTipoContainer.sDescricao = sDescricao
    objTipoContainer.dIso = dIso
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TipoContainer_Le = SUCESSO

    Exit Function

Erro_TipoContainer_Le:

    TipoContainer_Le = gErr

    Select Case gErr

        Case 97083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97084, 97085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER", gErr, objTipoContainer.iTipo)

        Case 97086 'nao achou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoContainer_Le_Todos(colItens As Collection) As Long
'le todos os tipos de container que estao no bd

Dim lComando As Long
Dim objTipoContainer As ClassTipoContainer
Dim iCod As Integer
Dim sDesc As String
Dim lErro As Long
Dim dIso As Double

On Error GoTo Erro_TipoContainer_Le_Todos

    'abre conexao com o bd
    lComando = Comando_Abrir
    If lComando = 0 Then gError 97087

    'seleciona os tipos de container
    sDesc = String(STRING_TIPOCONTAINER_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT tipo, descricao,ISO FROM TipoContainer order by Tipo", iCod, sDesc, dIso)
    If lErro <> AD_SQL_SUCESSO Then gError 97088

    'busca o primeiro tipo de container da resultado da operacao anterior...
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97089
    
    If lErro = AD_SQL_SEM_DADOS Then gError 97091
    
    'vai lendo os itens ate que leia todos os selecionados
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'carrega no obj os dados do item
        Set objTipoContainer = New ClassTipoContainer
        
        objTipoContainer.iTipo = iCod
        objTipoContainer.sDescricao = sDesc
        objTipoContainer.dIso = dIso
        
        'adiciona na colecao o obj...
        colItens.Add objTipoContainer
        
        'busca o proximo selecionado...
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97090
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    TipoContainer_Le_Todos = SUCESSO
    
    Exit Function

Erro_TipoContainer_Le_Todos:

    TipoContainer_Le_Todos = gErr
    
    Select Case gErr
    
        Case 97087
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case 97088 To 97090
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER_TODOS", gErr)
        
        Case 97091
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function OrigemDestino_Le(objOrigemDestino As ClassOrigemDestino) As Long
'Lê a OrigemDestino relacionada ao código passado no objOrigemDestino

Dim lErro As Long
Dim lComando As Long
Dim sOrigemDestino As String
Dim sUF As String

On Error GoTo Erro_OrigemDestino_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96564
        
    sOrigemDestino = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    sUF = String(STRING_ORIGEMDESTINO_UF, 0)
    
    'Verifica se existe no BD a OrigemDestino com o Código passado
    lErro = Comando_Executar(lComando, "SELECT OrigemDestino, UF FROM OrigemDestino WHERE Codigo = ?", sOrigemDestino, sUF, objOrigemDestino.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96565

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96566

    'Se o Código não existe no BD.
    If lErro = AD_SQL_SEM_DADOS Then gError 96567

    objOrigemDestino.sOrigemDestino = sOrigemDestino
    objOrigemDestino.sUF = sUF
    
    Call Comando_Fechar(lComando)

    OrigemDestino_Le = SUCESSO

    Exit Function

Erro_OrigemDestino_Le:

    OrigemDestino_Le = gErr

    Select Case gErr

        Case 96564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96565, 96566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)

        Case 96567
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProgNavio_Le(objProgNavio As ClassProgNavio) As Long
'Lê a Programação do Navio relacionada ao código passado no objProgNavio

Dim lErro As Long
Dim lComando As Long
Dim tProgNavio As TypeProgNavio

On Error GoTo Erro_ProgNavio_Le

    lComando = Comando_Abrir
    If lComando = 0 Then gError 96654
        
    With tProgNavio
    
        .sNavio = String(STRING_PROGNAVIO_NAVIO, 0)
        .sTerminal = String(STRING_PROGNAVIO_TERMINAL, 0)
        .sArmador = String(STRING_PROGNAVIO_ARMADOR, 0)
        .sAgMaritima = String(STRING_PROGNAVIO_AGMARITIMA, 0)
        .sViagem = String(STRING_PROGNAVIO_VIAGEM, 0)
        .sObservacao = String(STRING_PROGNAVIO_OBSERVACAO, 0)
        
        'Verifica se existe no BD a OrigemDestino com o Código passado
        lErro = Comando_Executar(lComando, "SELECT Navio, Terminal, Armador, AgMaritima, Viagem, Observacao, DataChegada, HoraChegada, DataDeadLine, HoraDeadLine FROM ProgNavio WHERE Codigo = ?", .sNavio, .sTerminal, .sArmador, .sAgMaritima, .sViagem, .sObservacao, .dtDataChegada, .dHoraChegada, .dtDataDeadLine, .dHoradeadLine, objProgNavio.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 96655
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96656
    
        'Se o Código não existe no BD.
        If lErro = AD_SQL_SEM_DADOS Then gError 96657
        
        objProgNavio.sNavio = .sNavio
        objProgNavio.sTerminal = .sTerminal
        objProgNavio.sArmador = .sArmador
        objProgNavio.sAgMaritima = .sAgMaritima
        objProgNavio.sViagem = .sViagem
        objProgNavio.sObservacao = .sObservacao
        objProgNavio.dtDataDeadLine = .dtDataDeadLine
        objProgNavio.dtDataChegada = .dtDataChegada
        objProgNavio.dtHoraChegada = .dHoraChegada
        objProgNavio.dtHoraDeadLine = .dHoradeadLine
    
    End With
    
    Call Comando_Fechar(lComando)

    ProgNavio_Le = SUCESSO

    Exit Function

Erro_ProgNavio_Le:

    ProgNavio_Le = gErr

    Select Case gErr

        Case 96654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96655, 96656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96657
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Despachante_Le(objDespachante As ClassDespachante) As Long
'Lê os dados da tabela Despachante cujo código foi passado.
'Devolve os dados em objDespachante.

Dim lErro As Long
Dim lComando As Long
Dim tDespachante As typeDespachante

On Error GoTo Erro_Despachante_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96676
    
    'Inicializa as variáveis
    tDespachante.sCGC = String(STRING_DESPACHANTE_CGC, 0)
    tDespachante.sNome = String(STRING_DESPACHANTE_NOME, 0)
    tDespachante.sNomeReduzido = String(STRING_DESPACHANTE_NOMEREDUZIDO, 0)

    'Lê os dados da tabela Despachante com o código passado.
    lErro = Comando_Executar(lComando, "SELECT CGC, Nome, NomeReduzido, Endereco FROM Despachante WHERE Codigo = ? " _
    , tDespachante.sCGC, tDespachante.sNome, tDespachante.sNomeReduzido, tDespachante.lEndereco, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96677

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96678
    
    ' Se não existe Despachante com o Código passado --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 96679

    'Carrega os dados em objDespachante
    objDespachante.sCGC = tDespachante.sCGC
    objDespachante.sNome = tDespachante.sNome
    objDespachante.sNomeReduzido = tDespachante.sNomeReduzido
    objDespachante.lEndereco = tDespachante.lEndereco
        
    Call Comando_Fechar(lComando)

    Despachante_Le = SUCESSO

    Exit Function

Erro_Despachante_Le:

    Despachante_Le = gErr

    Select Case gErr

        Case 96676
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96677, 96678
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPACHANTE", gErr, objDespachante.iCodigo)

        Case 96679
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function DespachanteContatos_Le(objDespachante As ClassDespachante) As Long
'Lê todos os Contatos associados ao Despachante cujo código foi passado
'Devolve em objCondPagtoAdm.colCondPagtoLoja

Dim lErro As Long
Dim lComando As Long
Dim objContato As ClassContato
Dim sContato As String
Dim sFax As String
Dim sEmail As String
Dim sTelefone As String
Dim sSetor As String

On Error GoTo Erro_DespachanteContatos_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96680
    
    'Inicializa as variáveis
    sContato = String(STRING_CONTATO_CONTATO, 0)
    sFax = String(STRING_CONTATO_FAX, 0)
    sEmail = String(STRING_CONTATO_EMAIL, 0)
    sTelefone = String(STRING_CONTATO_TELEFONE, 0)
    sSetor = String(STRING_CONTATO_SETOR, 0)
    
    'Lê os dados da tabela Contato com o código do despachante passado
    lErro = Comando_Executar(lComando, "SELECT Contato, Fax, Email, Telefone, Setor FROM Contato WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", _
    sContato, sFax, sEmail, sTelefone, sSetor, DESPACHANTE, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96681

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96682
         
    'Se não existir --> erro.
    Do While lErro = AD_SQL_SUCESSO
    
        'Inicializa contato
        Set objContato = New ClassContato
        
        'Carrega os dados em objContato
        objContato.sContato = sContato
        objContato.sEmail = sEmail
        objContato.sFax = sFax
        objContato.sSetor = sSetor
        objContato.sTelefone = sTelefone
        
        'Preenche na coleção
        objDespachante.colContato.Add objContato

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96683
        
    Loop

    Call Comando_Fechar(lComando)

    DespachanteContatos_Le = SUCESSO

    Exit Function

Erro_DespachanteContatos_Le:

    DespachanteContatos_Le = gErr

    Select Case gErr

        Case 96680
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96681, 96682, 96683
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabPrecoAntData_Le(objTabPreco As ClassTabPreco) As Long
'Lê os dados da tabela TabPreco cujo código foi passado e a data é a Data Atual.

Dim lErro As Long
Dim lComando As Long
Dim tTabPreco As typeTabPreco

On Error GoTo Erro_TabPrecoAntData_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96781
           
    With tTabPreco
    
        'Inicializa
        .sObservacao = String(STRING_TABPRECO_OBSERVACAO, 0)
    
        'Lê os dados da tabela TabPreco associados ao código e menores ou igual a Data Atual.
        lErro = Comando_Executar(lComando, "SELECT DataVigencia, AdValoren, Pedagio, Destino, Origem, Cliente, Observacao FROM TabPrecoGR WHERE Codigo = ? and DataVigencia <= ? ORDER BY DataVigencia desc" _
        , .dtDataVigencia, .dAdValoren, .dPedagio, .iDestino, .iOrigem, .lCliente, .sObservacao, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 96782
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96783
    
        'Se não existe Tabela de Preço com o Código e a data passada --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 96784
        
        'Carrega os dados em objTabPreco
        objTabPreco.lCliente = .lCliente
        objTabPreco.dAdValoren = .dAdValoren
        objTabPreco.dPedagio = .dPedagio
        objTabPreco.iDestino = .iDestino
        objTabPreco.iOrigem = .iOrigem
        objTabPreco.sObservacao = .sObservacao
        objTabPreco.dtDataVigencia = .dtDataVigencia
    
    End With
    
    Call Comando_Fechar(lComando)

    TabPrecoAntData_Le = SUCESSO

    Exit Function

Erro_TabPrecoAntData_Le:

    TabPrecoAntData_Le = gErr

    Select Case gErr

        Case 96781
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96782, 96783
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)

        Case 96784
            'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrigemDestino_Le_UF(objOrigemDestino As ClassOrigemDestino) As Long
'Lê a OrigemDestino a UF passados no objOrigemDestino

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_OrigemDestino_Le_UF

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99274
    
    'Verifica se existe no BD a OrigemDestino com o Nome e a UF passados
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM OrigemDestino WHERE UF = ?", iCodigo, objOrigemDestino.sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 99275

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99276

    'Se o Código não existe no BD.
    If lErro = AD_SQL_SEM_DADOS Then gError 99277
    
    'Carrega o obj com o código encontrado
    objOrigemDestino.iCodigo = iCodigo
    
    Call Comando_Fechar(lComando)

    OrigemDestino_Le_UF = SUCESSO

    Exit Function

Erro_OrigemDestino_Le_UF:

    OrigemDestino_Le_UF = gErr

    Select Case gErr

        Case 99274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99275, 99276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO_UF", gErr, objOrigemDestino.sUF)

        Case 99277
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
Function OrigemDestino_Le_NomeUF(objOrigemDestino As ClassOrigemDestino) As Long
'Lê a OrigemDestino relacionada ao Nome e a UF passados no objOrigemDestino

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_OrigemDestino_Le_NomeUF

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96857
    
    'Verifica se existe no BD a OrigemDestino com o Nome e a UF passados
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM OrigemDestino WHERE OrigemDestino = ? and UF = ?", iCodigo, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 96858

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96859

    'Se o Código não existe no BD.
    If lErro = AD_SQL_SEM_DADOS Then gError 96860
    
    'Carrega o obj com o código encontrado
    objOrigemDestino.iCodigo = iCodigo
    
    Call Comando_Fechar(lComando)

    OrigemDestino_Le_NomeUF = SUCESSO

    Exit Function

Erro_OrigemDestino_Le_NomeUF:

    OrigemDestino_Le_NomeUF = gErr

    Select Case gErr

        Case 96857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96858, 96859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO_NOMEUF", gErr, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)

        Case 96860
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrigemDestino_Le_CodigoUF(objOrigemDestino As ClassOrigemDestino) As Long
'Lê a OrigemDestino relacionada ao Código e a UF passados no objOrigemDestino

Dim lErro As Long
Dim lComando As Long
Dim sNome As String

On Error GoTo Erro_OrigemDestino_Le_CodigoUF

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96865
    
    sNome = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    
    'Verifica se existe no BD a OrigemDestino com o Código e a UF passados
    lErro = Comando_Executar(lComando, "SELECT OrigemDestino FROM OrigemDestino WHERE Codigo = ? and UF = ?", sNome, objOrigemDestino.iCodigo, objOrigemDestino.sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 96866

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96867

    'Se a OrigemDestino não existe no BD.
    If lErro = AD_SQL_SEM_DADOS Then gError 96868
    
    'Carrega o obj com o nome encontrado
    objOrigemDestino.sOrigemDestino = sNome
    
    Call Comando_Fechar(lComando)

    OrigemDestino_Le_CodigoUF = SUCESSO

    Exit Function

Erro_OrigemDestino_Le_CodigoUF:

    OrigemDestino_Le_CodigoUF = gErr

    Select Case gErr

        Case 96865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96866, 96867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO_CODIGOUF", gErr, objOrigemDestino.iCodigo, objOrigemDestino.sUF)

        Case 96868
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrigemDestino_Le_Nome(objOrigemDestino As ClassOrigemDestino) As Long
'Lê a OrigemDestino relacionada ao Nome passado no objOrigemDestino

Dim lErro As Long
Dim lComando As Long
Dim sUF As String
Dim iCodigo As Integer

On Error GoTo Erro_OrigemDestino_Le_Nome

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96861
    
    sUF = String(STRING_ORIGEMDESTINO_UF, 0)
    
    'Verifica se existe no BD a OrigemDestino com o Nome passado
    lErro = Comando_Executar(lComando, "SELECT UF, Codigo FROM OrigemDestino WHERE OrigemDestino = ? ", sUF, iCodigo, objOrigemDestino.sOrigemDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 96862

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96863

    'Se o Nome não existe no BD.
    If lErro = AD_SQL_SEM_DADOS Then gError 96864
    
    'Carrega o obj com os dados recolhidos no banco
    objOrigemDestino.sUF = sUF
    objOrigemDestino.iCodigo = iCodigo
    
    Call Comando_Fechar(lComando)

    OrigemDestino_Le_Nome = SUCESSO

    Exit Function

Erro_OrigemDestino_Le_Nome:

    OrigemDestino_Le_Nome = gErr

    Select Case gErr

        Case 96861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96862, 96863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO_NOME", gErr, objOrigemDestino.sOrigemDestino)

        Case 96864
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabPreco_Le(objTabPreco As ClassTabPreco) As Long
'Lê os dados da tabela TabPreco cujo código e Data de Vigência foram passadas.

Dim lErro As Long
Dim lComando As Long
Dim tTabPreco As typeTabPreco

On Error GoTo Erro_TabPreco_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96768
     
    With tTabPreco
    
        'Inicializa
        .sObservacao = String(STRING_TABPRECO_OBSERVACAO, 0)
    
        'Lê os dados da tabela TabPreco.
        lErro = Comando_Executar(lComando, "SELECT AdValoren, Pedagio, Destino, Origem, Cliente, Observacao FROM TabPrecoGR WHERE Codigo = ? and DataVigencia = ?" _
        , .dAdValoren, .dPedagio, .iDestino, .iOrigem, .lCliente, .sObservacao, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 96769
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96770
    
        'Se não existe Tabela de Preço com o Código e a data passada --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 96771
    
        'Carrega os dados em objTabPreco
        objTabPreco.lCliente = .lCliente
        objTabPreco.dAdValoren = .dAdValoren
        objTabPreco.dPedagio = .dPedagio
        objTabPreco.iDestino = .iDestino
        objTabPreco.iOrigem = .iOrigem
        objTabPreco.sObservacao = .sObservacao
    
    End With
    
    Call Comando_Fechar(lComando)

    TabPreco_Le = SUCESSO

    Exit Function

Erro_TabPreco_Le:

    TabPreco_Le = gErr

    Select Case gErr

        Case 96768
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96769, 96770
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)

        Case 96771
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Verifica_Servico_TabPreco(ByVal objTabPreco As ClassTabPreco) As Long
'Verifica cada produto da tabela de preço para ver se está relacionado a alguma solicitação
'Verifica se o serviço deve ser excluído ou incluído e adiciona em coleções

Dim lErro As Long
Dim ColServicoSol As New Collection
Dim ColServicoBD As New Collection
Dim colservicoExclui As New Collection
Dim colservicoInclui As New Collection

On Error GoTo Erro_Verifica_Servico_TabPreco

    'Preenche a coleção com os dados do serviço existente na solicitação
    lErro = CF("Preenche_ColServicoSol1", ColServicoSol, objTabPreco)
    If lErro <> SUCESSO Then gError 98356
    
    'Preenche a coleção com os dados do serviço existente nos itens da tabela de preço
    lErro = CF("Preenche_ColServicoBD", ColServicoBD, objTabPreco)
    If lErro <> SUCESSO Then gError 98357
    
    'Compara os serviços do banco com os da tela para verificar os que devem ser incluídos
    lErro = CF("Testa_Servico_Inclui1", objTabPreco.colTabPrecoItens, ColServicoBD, colservicoInclui)
    If lErro <> SUCESSO Then gError 98358

    'Compara os serviços do banco com os da tela para verificar os que devem ser excluídos
    lErro = CF("Testa_Servico_Exclui1", objTabPreco.colTabPrecoItens, ColServicoSol, ColServicoBD, colservicoExclui)
    If lErro <> SUCESSO Then gError 98359
    
    'Exclui os dados do serviço já verificados
    lErro = CF("TabPrecoItens_Exclui", colservicoExclui, objTabPreco)
    If lErro <> SUCESSO Then gError 98360

    'Inclui os novos dados do serviço já verificados
    lErro = CF("TabPrecoItens_Inclui", colservicoInclui)
    If lErro <> SUCESSO Then gError 98361

    Verifica_Servico_TabPreco = SUCESSO

    Exit Function

Erro_Verifica_Servico_TabPreco:

    Verifica_Servico_TabPreco = gErr

    Select Case gErr

        Case 98356 To 98362

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Testa_Servico_Exclui1(ByVal ColServico As Collection, ByVal ColServicoSol As Collection, ByVal ColServicoBD As Collection, ByVal colservicoExclui As Collection) As Long
'Compara os serviços do banco com os da tela para verificar os que devem ser excluídos

Dim bNaoAchou_Tela As Boolean
Dim objServicoBD As New ClassServico
Dim objServico As New ClassTabPrecoItens
Dim bNaoAchou_Sol As Boolean
Dim lErro As Long

On Error GoTo Erro_Testa_Servico_Exclui1
   
    'Descobrir os que devem ser excluidos
    For Each objServicoBD In ColServicoBD

        For Each objServico In ColServico
        
             bNaoAchou_Tela = True
            'Verificar se o serviço existe na tela
            If objServicoBD.sProduto = objServico.sProduto Then
                bNaoAchou_Tela = False
                Exit For
            End If
        Next

        'Se existe no banco e não existe na tela...
        If bNaoAchou_Tela Then

            'Verificar se existe a solicitação associada ao serviço
            lErro = Testa_Existencia_Solicitacao(ColServicoSol, objServicoBD, bNaoAchou_Sol)
            If lErro <> SUCESSO Then gError 98366

            'se não existe Solicitação relacionada ao serviço--> exclui
            If bNaoAchou_Sol Then
                colservicoExclui.Add objServicoBD
            'se existe comprovante relacionado ao serviço--> erro
            Else
                gError 98367
            End If

        End If

    Next

    Testa_Servico_Exclui1 = SUCESSO

    Exit Function

Erro_Testa_Servico_Exclui1:

    Testa_Servico_Exclui1 = gErr

    Select Case gErr

        Case 98366

        Case 98367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SERVICO1", gErr, objServicoBD.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Testa_Servico_Inclui1(ByVal ColServico As Collection, ByVal ColServicoBD As Collection, ByVal colservicoInclui As Collection) As Long
'Compara os serviços do banco com os da tela para verificar os que devem ser incluídos

Dim bNaoAchou_BD As Boolean
Dim objServicoBD As New ClassServico
Dim objServico As New ClassTabPrecoItens

On Error GoTo Erro_Testa_Servico_Inclui1
   
    'Descobrir os serviços que devem ser incluídos
    For Each objServico In ColServico
        For Each objServicoBD In ColServicoBD
        
             bNaoAchou_BD = True
            'Se existe no banco, sai.
            If objServicoBD.sProduto = objServico.sProduto Then
                bNaoAchou_BD = False
                Exit For
            End If
            
        Next
        'Se não existe --> Inclui
        If bNaoAchou_BD Then colservicoInclui.Add objServico
    Next

    Testa_Servico_Inclui1 = SUCESSO

    Exit Function

Erro_Testa_Servico_Inclui1:

    Testa_Servico_Inclui1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Testa_Existencia_Solicitacao(ByVal ColServicoSol As Collection, ByVal objServicoBD As ClassServico, bNaoAchou_Sol As Boolean) As Long
'Verifica se o serviço da Tabela de Preço está associado a alguma solicitação

Dim objServicoSol As New ClassServico

On Error GoTo Erro_Testa_Existencia_Solicitacao

    bNaoAchou_Sol = True

    For Each objServicoSol In ColServicoSol

        'Se o serviço do banco estiver sendo usado por alguma solicitação
        If objServicoSol.sProduto = objServicoBD.sProduto Then

            bNaoAchou_Sol = False
            Exit For

        End If

    Next

    Testa_Existencia_Solicitacao = SUCESSO

    Exit Function

Erro_Testa_Existencia_Solicitacao:

    Testa_Existencia_Solicitacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function Preenche_ColServicoSol1(ByVal ColServicoSol As Collection, ByVal objTabPreco As ClassTabPreco) As Long
'Preenche a coleção com os dados do serviço existente na solicitação

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim objServicoSol As ClassServico

On Error GoTo Erro_Preenche_ColServicoSol1

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98352

    'Inicializa a variável
    sProduto = String(STRING_PRODUTO_CODIGO, 0)

    'Ler todos os servicos de SolServServico e colocar numa colecao
    lErro = Comando_Executar(lComando, "SELECT Produto FROM SolServServico, SolicitacaoServico WHERE CodTabPreco=? AND SolicitacaoServico.NumIntDoc=SolServServico.NumIntDocOrigem", sProduto, objTabPreco.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98353

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98354

    Do While lErro = AD_SQL_SUCESSO

        Set objServicoSol = New ClassServico

        objServicoSol.sProduto = sProduto
        
        'Preenche a coleção com o obj carregado
        ColServicoSol.Add objServicoSol

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98355

    Loop

    'Fecha o comando
    lComando = Comando_Fechar(lComando)

    Preenche_ColServicoSol1 = SUCESSO

    Exit Function

Erro_Preenche_ColServicoSol1:

    Preenche_ColServicoSol1 = gErr

    Select Case gErr

        Case 98352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98353, 98354, 98355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o comando
    lComando = Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Preenche_ColServicoBD(ByVal ColServicoBD As Collection, ByVal objTabPreco As ClassTabPreco) As Long
'Preenche a coleção com os dados do serviço existente nos itens da tabela de preço

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim objServicoBD As ClassServico

On Error GoTo Erro_Preenche_ColServicoBD

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98362

    'Inicializa a variável
    sProduto = String(STRING_PRODUTO_CODIGO, 0)

    'Ler todos os servicos da Tabela de Preço associados ao Código passado.
    lErro = Comando_Executar(lComando, "SELECT Produto FROM TabPrecoItensGR WHERE CodTabela=? AND DataVigencia=?", sProduto, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 98363

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98364

    Do While lErro = AD_SQL_SUCESSO

        Set objServicoBD = New ClassServico

        objServicoBD.sProduto = sProduto

        'Preenche a coleção com o obj carregado
        ColServicoBD.Add objServicoBD

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98365

    Loop

    'Fecha o comando
    lComando = Comando_Fechar(lComando)

    Preenche_ColServicoBD = SUCESSO

    Exit Function

Erro_Preenche_ColServicoBD:

    Preenche_ColServicoBD = gErr

    Select Case gErr

        Case 98363, 98364, 98365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)

        Case 98362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o comando
    lComando = Comando_Fechar(lComando)

    Exit Function

End Function

Function TabPrecoItens_Le(objTabPreco As ClassTabPreco) As Long
'Lê os dados da tabela TabPreco cujo código e Data de Vigência foram passadas.

Dim lErro As Long
Dim lComando As Long
Dim tTabPrecoItens As typeTabPrecoItens
Dim sDescricao As String
Dim objTabPrecoItens As ClassTabPrecoItens
Dim sCodigo As String

On Error GoTo Erro_TabPrecoItens_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96979
      
    With tTabPrecoItens
        
        'Inicializa as variáveis
        .sProduto = String(STRING_TABPRECOITENS_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTOS_DESCRICAO, 0)
    
        'Lê os produtos da tabela TabPrecoItens associados a tabela de preço na data de vigência.
        lErro = Comando_Executar(lComando, "SELECT Produto, Preco, Descricao FROM TabPrecoItensGR, Produtos WHERE CodTabela = ? and DataVigencia = ? AND TabPrecoItensGR.Produto = Produtos.Codigo" _
        , .sProduto, .dPreco, .sDescricao, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 96772
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96773
    
        'Se não existe Tabela de Preço de itens com o Código e a data passada --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 96774
    
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'carrega no obj os dados do item
            Set objTabPrecoItens = New ClassTabPrecoItens
    
            objTabPrecoItens.dPreco = .dPreco
            objTabPrecoItens.sDescricao = .sDescricao
            objTabPrecoItens.sProduto = .sProduto
    
            'adiciona na colecao o obj...
            objTabPreco.colTabPrecoItens.Add objTabPrecoItens
    
            'busca o proximo selecionado...
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97778
    
        Loop
        
    End With
    
    Call Comando_Fechar(lComando)
    
    TabPrecoItens_Le = SUCESSO

    Exit Function

Erro_TabPrecoItens_Le:

    TabPrecoItens_Le = gErr

    Select Case gErr

        Case 96772, 96773, 96778
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objTabPrecoItens.lCodTabela)

        Case 96774
            'Erro tratado na rotina chamadora
        
        Case 96979
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabPrecoItens_Le1(objTabPreco As ClassTabPreco) As Long
'Usado para verificar na hora da exclusão se existe itens associados
'a esta tabela de preço
Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_TabPrecoItens_Le1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98002
            
    'Verifica se existe uma tabela de preço com código e datade vigência passados
    lErro = Comando_Executar(lComando, "SELECT CodTabela FROM TabPrecoItensGR WHERE CodTabela= ? and DataVigencia = ?", lCodigo, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 98003

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98004

    'Se não está cadastrado --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98005
    
    Call Comando_Fechar(lComando)
    
    TabPrecoItens_Le1 = SUCESSO
    
    Exit Function
    
Erro_TabPrecoItens_Le1:

    TabPrecoItens_Le1 = gErr
    
    Select Case gErr
        
        Case 98002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98003, 98004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objTabPreco.lCodigo)
        
        Case 98005
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Documento_Le(objDocumento As ClassDocumento) As Long
'Lê os dados da tabela Documento cujo código foi passado.

Dim lErro As Long
Dim lComando As Long
Dim tDocumento As typeDocumento

On Error GoTo Erro_Documento_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98023
     
    With tDocumento
    
        'Inicializa
        .sDescricao = String(STRING_DOCUMENTO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_DOCUMENTO_NOMEREDUZIDO, 0)
        .sDocumento = String(STRING_DOCUMENTO_DOCUMENTO, 0)
        
        'Lê os dados do Documento.
        lErro = Comando_Executar(lComando, "SELECT Descricao, NomeReduzido, TipoDoc, Documento FROM Documento WHERE Codigo = ? ", .sDescricao, .sNomeReduzido, .iTipoDoc, .sDocumento, objDocumento.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 98024
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98025
    
        'Se não existe Documento com o Código passado --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 98026
    
        'Carrega os dados em objDocumento
        objDocumento.iTipoDoc = .iTipoDoc
        objDocumento.sDescricao = .sDescricao
        objDocumento.sNomeReduzido = .sNomeReduzido
        objDocumento.sDocumento = .sDocumento
        
    End With
    
    Call Comando_Fechar(lComando)

    Documento_Le = SUCESSO

    Exit Function

Erro_Documento_Le:

    Documento_Le = gErr

    Select Case gErr

        Case 98023
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98024, 98025
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO", gErr, objDocumento.iCodigo)

        Case 98026
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_CondPagto_LeLock(ByVal objCotacao As ClassCotacaoGR) As Long
'Le e faz o lock da Condicao de pagamento relacionada com a cotacao em questao

Dim lErro As Long
Dim lComando As Long
Dim iCod As Integer

On Error GoTo Erro_Cotacao_CondPagto_LeLock

    'Abre o  comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97376
     
    'Chama a funcao que fara o lock efetivo da CondPagto
    lErro = CF("CondicaoPagto_Lock", lComando, objCotacao.iCondicaoPagto)
    If lErro <> SUCESSO Then gError 97256
     
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
        
    Cotacao_CondPagto_LeLock = SUCESSO
    
    Exit Function

Erro_Cotacao_CondPagto_LeLock:

    Cotacao_CondPagto_LeLock = gErr
    
    Select Case gErr
    
        Case 97256
        
        Case 97376
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cotacao_Vendedor_LeLock(ByVal objCotacao As ClassCotacaoGR) As Long
'Le e faz o lock do Vendedor relacionado com a cotacao em questao

Dim lErro As Long
Dim lComando As Long
Dim iCod As Integer

On Error GoTo Erro_Cotacao_Vendedor_LeLock

    'Abre o  comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97377

    'Chama a funcao que fara o lock efetivo do vendedor
    lErro = CF("Vendedor_Lock", lComando, objCotacao.iCodVendedor)
    If lErro <> SUCESSO Then gError 97260
            
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Vendedor_LeLock = SUCESSO
    
    Exit Function

Erro_Cotacao_Vendedor_LeLock:

    Cotacao_Vendedor_LeLock = gErr
    
    Select Case gErr
    
        Case 97260
        
        Case 97337
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cotacao_TipoEmbalagem_LeLock(ByVal objCotacao As ClassCotacaoGR) As Long
'Le e faz o lock do TipoEmbalagem relacionado com a cotacao em questao

Dim lErro As Long
Dim lComando As Long
Dim iCod As Integer

On Error GoTo Erro_Cotacao_TipoEmbalagem_LeLock

    'Abre o  comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97378

    'Le o TipoEmbalagem
    lErro = Comando_ExecutarPos(lComando, "SELECT Tipo FROM TipoEmbalagem WHERE Tipo=? ", 0, iCod, objCotacao.iTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 97264
    
    'Posiciona no TipoEmbalagem
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97265
    
    If lErro <> AD_SQL_SUCESSO Then gError 97266
    
    'Lock do TipoEmbalagem
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 97267
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cotacao_TipoEmbalagem_LeLock = SUCESSO
    
    Exit Function

Erro_Cotacao_TipoEmbalagem_LeLock:

    Cotacao_TipoEmbalagem_LeLock = gErr
    
    Select Case gErr
    
        Case 97264, 97265
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM", gErr, objCotacao.iTipoEmbalagem)
        
        Case 97266
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_NAO_CADASTRADA", gErr, objCotacao.iTipoEmbalagem)
                        
        Case 97267
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOEMBALAGEM", gErr, objCotacao.iTipoEmbalagem)

        Case 97378
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cotacao_CotacaoContainer_LeLock(ByVal objCotacao As ClassCotacaoGR) As Long
'Le e faz o lock dos Containers relacionado com a cotacao em questao

Dim lErro As Long
Dim lComando As Long
Dim iCod As Integer
Dim objCotacaoContainer As ClassCotacaoContainer

On Error GoTo Erro_Cotacao_CotacaoContainer_LeLock

    'Abre o  comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97379
        
    For Each objCotacaoContainer In objCotacao.colCotacaoContainer
    
        'Le o TipoContainer
        lErro = Comando_ExecutarPos(lComando, "SELECT Tipo FROM TipoContainer WHERE Tipo=? ", 0, iCod, objCotacaoContainer.iTipoContainer)
        If lErro <> AD_SQL_SUCESSO Then gError 97268
        
        'Posiciona no TipoContainer
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97269
        
        If lErro <> AD_SQL_SUCESSO Then gError 97270
        
        'Lock do TipoContainer
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 97271
    
    Next
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cotacao_CotacaoContainer_LeLock = SUCESSO
    
    Exit Function

Erro_Cotacao_CotacaoContainer_LeLock:

    Cotacao_CotacaoContainer_LeLock = gErr
    
    Select Case gErr
    
        Case 97268, 97269
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER", gErr, objCotacaoContainer.iTipoContainer)
        
        Case 97270
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINER_NAO_CADASTRADO", gErr, objCotacaoContainer.iTipoContainer)
                        
        Case 97271
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOCONTAINER", gErr, objCotacaoContainer.iTipoContainer)

        Case 97379
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cotacao_CotacaoServico_LeLock(ByVal objCotacao As ClassCotacaoGR) As Long
'Le e faz o lock dos Servicos (produtos) relacionados com a cotacao em questao

Dim lErro As Long
Dim lComando As Long
Dim sCod As String
Dim objCotacaoServico As ClassCotacaoServico
Dim objProduto As New ClassProduto

On Error GoTo Erro_Cotacao_CotacaoServico_LeLock

    'Abre o  comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97380

    For Each objCotacaoServico In objCotacao.colCotacaoServico
                    
        objProduto.sCodigo = objCotacaoServico.sProduto

        lErro = CF("Produto_Lock_Shared", lComando, objProduto)
        If lErro <> SUCESSO Then gError 97290
    Next
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cotacao_CotacaoServico_LeLock = SUCESSO
    
    Exit Function

Erro_Cotacao_CotacaoServico_LeLock:

    Cotacao_CotacaoServico_LeLock = gErr
    
    Select Case gErr
    
        Case 97290

        Case 97380
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cotacao_LeGR(ByVal objCotacao As ClassCotacaoGR) As Long
'Lê a Cotacao a partir do Código e FilialEmpresa passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim tCotacao As typeCotacaoGR
Dim sObsDestOrigem1 As String
Dim sObsDestOrigem2 As String
Dim sObsDestOrigem3 As String
Dim sObsDestOrigem4 As String

On Error GoTo Erro_Cotacao_LeGR
   
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97166

    'inicializando
    tCotacao.sCliente = String(STRING_COTACAO_CLIENTE, 0)
    tCotacao.sDescCargaSolta = String(STRING_COTACAO_DESCCARGASOLTA, 0)
    tCotacao.sEnvioComplemento = String(STRING_COTACAO_ENVIOCOMPLEMENTO, 0)
    tCotacao.sIndicacao = String(STRING_COTACAO_INDICACAO, 0)
    tCotacao.sObservacao = String(STRING_COTACAO_OBSERVACAO, 0)
    tCotacao.sObsResultado = String(STRING_COTACAO_OBSRESULTADO, 0)
    sObsDestOrigem1 = String(STRING_COTACAO_OBSDESTORIGEM, 0)
    sObsDestOrigem2 = String(STRING_COTACAO_OBSDESTORIGEM, 0)
    sObsDestOrigem3 = String(STRING_COTACAO_OBSDESTORIGEM, 0)
    sObsDestOrigem4 = String(STRING_COTACAO_OBSDESTORIGEM, 0)
    
    'Busca na tabela CotacaoVenda a Cotacao com Código e FilialEmpresa informados
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, TipoOperacao, Cliente, Envio, EnvioComplemento, CodVendedor, Indicacao," _
    & "Observacao,DataPrevInicio, TipoEmbalagem, Ajudantes, Carga, CargaPorConta, Descarga, DescargaPorConta, Ova, OvaPorConta," _
    & "Desova, DesovaPorConta, CargaSolta, DescCargaSolta, CondicaoPagto, Situacao, Justificativa, ObsResultado,ValorMercadoria, ObsDestOrigem1, ObsDestOrigem2, ObsDestOrigem3, ObsDestOrigem4 FROM CotacaoGR " _
    & "WHERE Codigo=? AND FilialEmpresa=?", tCotacao.lNumIntDoc, tCotacao.iFilialEmpresa, tCotacao.lCodigo, tCotacao.dtData, tCotacao.iTipoOperacao, _
    tCotacao.sCliente, tCotacao.iEnvio, tCotacao.sEnvioComplemento, tCotacao.iCodVendedor, tCotacao.sIndicacao, tCotacao.sObservacao, _
    tCotacao.dtDataPrevInicio, tCotacao.iTipoEmbalagem, tCotacao.iAjudantes, tCotacao.iCarga, tCotacao.iCargaPorConta, tCotacao.iDesCarga, tCotacao.iDesCargaPorConta, _
    tCotacao.iOva, tCotacao.iOvaPorConta, tCotacao.iDesova, tCotacao.iDesovaPorConta, tCotacao.iCargaSolta, tCotacao.sDescCargaSolta, tCotacao.iCondicaoPagto, tCotacao.iSituacao, _
    tCotacao.iJustificativa, tCotacao.sObsResultado, tCotacao.dValorMerc, sObsDestOrigem1, sObsDestOrigem2, sObsDestOrigem3, sObsDestOrigem4, objCotacao.lCodigo, objCotacao.iFilialEmpresa)

    If lErro <> AD_SQL_SUCESSO Then gError 97168

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97169

    'Se não encontrou a Cotacao ==> erro 97170
    If lErro = AD_SQL_SEM_DADOS Then gError 97170

    'Carrega o obj com o conteudo lido do bd
    objCotacao.dtData = tCotacao.dtData
    objCotacao.dtDataPrevInicio = tCotacao.dtDataPrevInicio
    objCotacao.iAjudantes = tCotacao.iAjudantes
    objCotacao.iCarga = tCotacao.iCarga
    objCotacao.iCargaPorConta = tCotacao.iCargaPorConta
    objCotacao.iCargaSolta = tCotacao.iCargaSolta
    objCotacao.iCodVendedor = tCotacao.iCodVendedor
    objCotacao.iCondicaoPagto = tCotacao.iCondicaoPagto
    objCotacao.iDesCarga = tCotacao.iDesCarga
    objCotacao.iDesCargaPorConta = tCotacao.iDesCargaPorConta
    objCotacao.iDesova = tCotacao.iDesova
    objCotacao.iDesovaPorConta = tCotacao.iDesovaPorConta
    objCotacao.iEnvio = tCotacao.iEnvio
    objCotacao.iFilialEmpresa = tCotacao.iFilialEmpresa
    objCotacao.iJustificativa = tCotacao.iJustificativa
    objCotacao.iOva = tCotacao.iOva
    objCotacao.iOvaPorConta = tCotacao.iOvaPorConta
    objCotacao.iSituacao = tCotacao.iSituacao
    objCotacao.iTipoEmbalagem = tCotacao.iTipoEmbalagem
    objCotacao.iTipoOperacao = tCotacao.iTipoOperacao
    objCotacao.lCodigo = tCotacao.lCodigo
    objCotacao.lNumIntDoc = tCotacao.lNumIntDoc
    objCotacao.sCliente = tCotacao.sCliente
    objCotacao.sDescCargaSolta = tCotacao.sDescCargaSolta
    objCotacao.sEnvioComplemento = tCotacao.sEnvioComplemento
    objCotacao.sIndicacao = tCotacao.sIndicacao
    objCotacao.sObsDestOrigem = (sObsDestOrigem1 & sObsDestOrigem2 & sObsDestOrigem3 & sObsDestOrigem4)
    objCotacao.sObservacao = tCotacao.sObservacao
    objCotacao.sObsResultado = tCotacao.sObsResultado
    objCotacao.dValorMerc = tCotacao.dValorMerc
    
    Call Comando_Fechar(lComando)

    Cotacao_LeGR = SUCESSO

    Exit Function

Erro_Cotacao_LeGR:

    Cotacao_LeGR = gErr

    Select Case gErr

        Case 97166
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97168, 97169
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO1", gErr)

        Case 97170
            'NAO ACHOU COTACAO !!

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CotacaoOrigemDestino_Le_Cotacao(ByVal objCotacao As ClassCotacaoGR, ByVal colCotacoesOrigemDestino As Collection) As Long
'Lê as CotacoesOrigemDestino a partir de uma Cotacao

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoOrigemDestino As typeCotacaoOrigemDestino
Dim objCotacaoOrigemDestino As ClassCotacaoOrigemDestino

On Error GoTo Erro_CotacaoOrigemDestino_Le_Cotacao

    'inicializando...
    tCotacaoOrigemDestino.sDestino = String(STRING_COTACAOORIGEMDESTINO_DESTINO, 0)
    tCotacaoOrigemDestino.sOrigem = String(STRING_COTACAOORIGEMDESTINO_ORIGEM, 0)
    tCotacaoOrigemDestino.sServico = String(STRING_COTACAOORIGEMDESTINO_SERVICO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97180

    'Busca na tabela CotacaoOrigemDestino as CotacoesOrigemDestino
    'relacionadas com a cotacao em questao
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, Servico, Origem, Destino from CotacaoOrigemDestino where NumIntDocOrigem = ?", tCotacaoOrigemDestino.lNumIntDoc, tCotacaoOrigemDestino.lNumIntDocOrigem, tCotacaoOrigemDestino.sServico, tCotacaoOrigemDestino.sOrigem, tCotacaoOrigemDestino.sDestino, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97181

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97182

    'Se não encontrou nenhuma CotacaoOrigemDestino ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97183

    'preenche a colecao com os dados antes lidos
    'um obj eh utilizado para que o armazenamento
    'seja feito
    Do While lErro = AD_SQL_SUCESSO
        
        Set objCotacaoOrigemDestino = New ClassCotacaoOrigemDestino
        
        objCotacaoOrigemDestino.lNumIntDoc = tCotacaoOrigemDestino.lNumIntDoc
        objCotacaoOrigemDestino.lNumIntDocOrigem = tCotacaoOrigemDestino.lNumIntDocOrigem
        objCotacaoOrigemDestino.sDestino = tCotacaoOrigemDestino.sDestino
        objCotacaoOrigemDestino.sOrigem = tCotacaoOrigemDestino.sOrigem
        objCotacaoOrigemDestino.sServico = tCotacaoOrigemDestino.sServico
        
        colCotacoesOrigemDestino.Add objCotacaoOrigemDestino

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97184

    Loop

    Call Comando_Fechar(lComando)

    CotacaoOrigemDestino_Le_Cotacao = SUCESSO

    Exit Function

Erro_CotacaoOrigemDestino_Le_Cotacao:

    CotacaoOrigemDestino_Le_Cotacao = gErr

    Select Case gErr

        Case 97180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97181, 97182, 97184
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOORIGEMDESTINO", gErr)

        Case 97183
            'Nao encontrou nenhuma CotacaoOrigemDestino relacionada a Cotacao passada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CotacaoServico_Le_Cotacao(ByVal objCotacao As ClassCotacaoGR, ByVal colCotacoesServico As Collection) As Long
'Lê as CotacoesServico a partir de uma Cotacao

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoServico As typeCotacaoServico
Dim objCotacaoServico As ClassCotacaoServico

On Error GoTo Erro_CotacaoServico_Le_Cotacao

    'incializando...
    tCotacaoServico.sProduto = String(STRING_COTACAOSERVICO_PRODUTO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97185

    'Busca na tabela CotacaoServico as CotacoesServico
    'relacionadas com a cotacao em questao
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, Produto, Quantidade, PrecoUnitario, AdValoren, Pedagio,CodigoOrigem,CodigoDestino from CotacaoServico where NumIntDocOrigem = ?", tCotacaoServico.lNumIntDoc, tCotacaoServico.lNumIntDocOrigem, tCotacaoServico.sProduto, tCotacaoServico.dQuantidade, tCotacaoServico.dPrecoUnitario, tCotacaoServico.dAdValoren, tCotacaoServico.dPedagio, tCotacaoServico.iOrigem, tCotacaoServico.iDestino, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97186

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97187

    'Se não encontrou nenhuma CotacaoContainer ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97188

    'preenche a colecao com os dados antes lidos
    'um obj eh utilizado para que o armazenamento
    'seja feito
    Do While lErro = AD_SQL_SUCESSO
        
        Set objCotacaoServico = New ClassCotacaoServico
        
        objCotacaoServico.dAdValoren = tCotacaoServico.dAdValoren
        objCotacaoServico.dPedagio = tCotacaoServico.dPedagio
        objCotacaoServico.dPrecoUnitario = tCotacaoServico.dPrecoUnitario
        objCotacaoServico.dQuantidade = tCotacaoServico.dQuantidade
        objCotacaoServico.lNumIntDoc = tCotacaoServico.lNumIntDoc
        objCotacaoServico.lNumIntDocOrigem = tCotacaoServico.lNumIntDocOrigem
        objCotacaoServico.sProduto = tCotacaoServico.sProduto
        objCotacaoServico.iDestino = tCotacaoServico.iDestino
        objCotacaoServico.iOrigem = tCotacaoServico.iOrigem
        
        colCotacoesServico.Add objCotacaoServico

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97189

    Loop

    Call Comando_Fechar(lComando)

    CotacaoServico_Le_Cotacao = SUCESSO

    Exit Function

Erro_CotacaoServico_Le_Cotacao:

    CotacaoServico_Le_Cotacao = gErr

    Select Case gErr

        Case 97185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97186, 97187, 97189
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOSERVICO", gErr)

        Case 97188
            'Nao encontrou nenhuma CotacaoServico relacionada a Cotacao passada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CotacaoContainer_Le_Cotacao(ByVal objCotacao As ClassCotacaoGR, ByVal colCotacoesContainer As Collection) As Long
'Lê as CotacoesContainer a partir de uma Cotacao

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoContainer As typeCotacaoContainer
Dim objCotacaoContainer As ClassCotacaoContainer

On Error GoTo Erro_CotacaoContainer_Le_Cotacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97172

    'Busca na tabela CotacaoContainer as CotacoesContainer
    'relacionadas com a cotacao em questao
    'e armazena no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, TipoContainer, Quantidade from CotacaoContainer where NumIntDocOrigem = ?", tCotacaoContainer.lNumIntDoc, tCotacaoContainer.lNumIntDocOrigem, tCotacaoContainer.iTipoContainer, tCotacaoContainer.iQuantidade, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97176

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97177

    'Se não encontrou nenhuma CotacaoContainer ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97178

    'preenche a colecao com os dados antes lidos
    'um obj eh utilizado para que o armazenamento
    'seja feito
    Do While lErro = AD_SQL_SUCESSO
    
        Set objCotacaoContainer = New ClassCotacaoContainer
        
        objCotacaoContainer.iQuantidade = tCotacaoContainer.iQuantidade
        objCotacaoContainer.iTipoContainer = tCotacaoContainer.iTipoContainer
        objCotacaoContainer.lNumIntDoc = tCotacaoContainer.lNumIntDoc
        objCotacaoContainer.lNumIntDocOrigem = tCotacaoContainer.lNumIntDocOrigem

        colCotacoesContainer.Add objCotacaoContainer

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97179

    Loop

    Call Comando_Fechar(lComando)

    CotacaoContainer_Le_Cotacao = SUCESSO

    Exit Function

Erro_CotacaoContainer_Le_Cotacao:

    CotacaoContainer_Le_Cotacao = gErr

    Select Case gErr

        Case 97172
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97176, 97177, 97179
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOCONTAINER", gErr)

        Case 97178
            'Nao encontrou nenhuma CotacaoContainer relacionada a Cotacao passada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Contato_Le_Cotacao(ByVal objCotacao As ClassCotacaoGR, ByVal colContatos As Collection) As Long
'Lê os Contatos a partir de uma Cotacao com tiponumintdocorigem = 2 (Cotacao)

Dim lErro As Long
Dim lComando As Long
Dim sContato As String
Dim sFax As String
Dim sEmail As String
Dim sTelefone As String
Dim sSetor As String
Dim objContato As ClassContato

On Error GoTo Erro_Contato_Le_Cotacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97191

    'Inicializa
    sContato = String(STRING_CONTATO_CONTATO, 0)
    sFax = String(STRING_CONTATO_FAX, 0)
    sEmail = String(STRING_CONTATO_EMAIL, 0)
    sTelefone = String(STRING_CONTATO_TELEFONE, 0)
    sSetor = String(STRING_CONTATO_SETOR, 0)
    
    'Lê os dados da tabela Contato
    lErro = Comando_Executar(lComando, "SELECT Contato, Fax, Email, Telefone, Setor FROM Contato WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", _
    sContato, sFax, sEmail, sTelefone, sSetor, TIPO_DOC_ORIGEM_COTACAO, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97192

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97193
         
    If lErro = AD_SQL_SEM_DADOS Then gError 97194
         
    Do While lErro = AD_SQL_SUCESSO
    
        'Inicializa contato
        Set objContato = New ClassContato
        
        'Carrega os dados em objContato
        objContato.sContato = sContato
        objContato.sEmail = sEmail
        objContato.sFax = sFax
        objContato.sSetor = sSetor
        objContato.sTelefone = sTelefone
        
        If Len(Trim(sTelefone)) = 7 Then
            objContato.sTelefone = Format(sTelefone, "###-####")
        ElseIf Len(Trim(sTelefone)) = 8 Then
            objContato.sTelefone = Format(sTelefone, "####-####")
        End If
        
        colContatos.Add objContato

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97195
        
    Loop

    Call Comando_Fechar(lComando)

    Contato_Le_Cotacao = SUCESSO

    Exit Function

Erro_Contato_Le_Cotacao:

    Contato_Le_Cotacao = gErr

    Select Case gErr

        Case 97191
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97192, 97193, 97195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATO", gErr)

        Case 97194
            'Nao encontrou nenhum Contato relacionado a Cotacao passada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Verifica_Servico(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Verifica se o serviço deve ser excluído ou incluído e adiciona em coleções

Dim lErro As Long
Dim ColServicoSol As New Collection
Dim ColServicoComp As New Collection
Dim colservicoExclui As New Collection
Dim colservicoInclui As New Collection

On Error GoTo Erro_Verifica_Servico
    
    'Preenche a coleção com os dados do serviço existente no banco
    lErro = CF("Preenche_ColServicoSol", ColServicoSol, objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98347
    
    'Preenche a coleção com os dados do serviço existente no Comprovante associado a este serviço
    lErro = CF("Preenche_ColServicoComp", ColServicoComp, objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98348
    
    'Descobre os serviços da solicitação que podem ser incluídos
    lErro = CF("Testa_Servico_Inclui", objSolicitacaoServico.ColServico, ColServicoSol, colservicoInclui)
    If lErro <> SUCESSO Then gError 98319
    
    'Descobre os serviços da solicitação que podem ser excluídos
    lErro = CF("Testa_Servico_Exclui", objSolicitacaoServico.ColServico, ColServicoSol, ColServicoComp, colservicoExclui)
    If lErro <> SUCESSO Then gError 98320
    
    'Descobre os serviços da solicitação que podem ser alterados
    lErro = CF("Testa_Servico_Altera", objSolicitacaoServico.ColServico, ColServicoSol, ColServicoComp, colservicoExclui, colservicoInclui)
    If lErro <> SUCESSO Then gError 98321
        
    'Exclui os dados do serviço já verificados
    lErro = CF("SolServServico_Exclui", colservicoExclui, objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98312
                                        
    'Inclui os novos dados do serviço já verificados
    lErro = CF("SolServServico_Inclui", colservicoInclui, objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98313
            
    Verifica_Servico = SUCESSO

    Exit Function

Erro_Verifica_Servico:

    Verifica_Servico = gErr

    Select Case gErr
                           
        Case 98312, 98313, 98319, 98320, 98321, 98347, 98348
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

Function Testa_Servico_Altera(ByVal ColServico As Collection, ByVal ColServicoSol As Collection, ByVal ColServicoComp As Collection, ByVal colservicoExclui As Collection, ByVal colservicoInclui As Collection) As Long
'Descobre os serviços da solicitação que podem ser alterados

Dim bAchou_Tela As Boolean
Dim bNaoAchou_Comp As Boolean
Dim objServicoSol As New ClassServico
Dim objServico As New ClassServico
Dim lErro As Long

On Error GoTo Erro_Testa_Servico_Altera

    bAchou_Tela = False
    
    'Descobrir os que devem ser alterados
    For Each objServicoSol In ColServicoSol
    
        For Each objServico In ColServico
            'Verificar se o serviço existe na tela com quantidade diferente
            If objServicoSol.sProduto = objServico.sProduto And objServicoSol.dQuant <> objServico.dQuant Then
                bAchou_Tela = True
                Exit For
            End If
        Next
        
        'Se existe na tela com quantidade diferente da anterior
        If bAchou_Tela Then
        
            'Verificar se existe comprovante de serviço associado
            lErro = CF("Testa_Existencia_Comprovante", ColServicoComp, objServicoSol, bNaoAchou_Comp)
            If lErro <> SUCESSO Then gError 98310
            
            'se não existe comprovante relacionado ao serviço--> altera
            If bNaoAchou_Comp Then
                colservicoExclui.Add objServicoSol
                colservicoInclui.Add objServico
                
            'se existe comprovante relacionado ao serviço--> erro
            Else
                gError 98323
                
            End If
           
        End If
        
    Next
        
    Testa_Servico_Altera = SUCESSO
    
    Exit Function
    
Erro_Testa_Servico_Altera:
    
    Testa_Servico_Altera = gErr
    
    Select Case gErr
        
        Case 98310
        
        Case 98323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SERVICO", gErr, objServicoSol.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
     
End Function

Function Testa_Servico_Exclui(ByVal ColServico As Collection, ByVal ColServicoSol As Collection, ByVal ColServicoComp As Collection, ByVal colservicoExclui As Collection) As Long
'Descobre os serviços da solicitação que podem ser excluídos

Dim bNaoAchou_Tela As Boolean
Dim objServicoSol As New ClassServico
Dim objServico As New ClassServico
Dim bNaoAchou_Comp As Boolean
Dim lErro As Long

On Error GoTo Erro_Testa_Servico_Exclui
   
    'Descobrir os que devem ser excluidos
    For Each objServicoSol In ColServicoSol
        For Each objServico In ColServico
        
            bNaoAchou_Tela = True
            'Verificar se o serviço existe na tela
            If objServicoSol.sProduto = objServico.sProduto Then
                bNaoAchou_Tela = False
                Exit For
            End If
            
        Next
        
        'Se existe no banco e não existe na tela...
        If bNaoAchou_Tela Then
            
            'Verificar se existe comprovante de serviço associado
            lErro = CF("Testa_Existencia_Comprovante", ColServicoComp, objServicoSol, bNaoAchou_Comp)
            If lErro <> SUCESSO Then gError 98311
            
            'se não existe comprovante relacionado ao serviço--> exclui
            If bNaoAchou_Comp Then
                colservicoExclui.Add objServicoSol
            'se existe comprovante relacionado ao serviço--> erro
            Else
                gError 98322
                
            End If
            
        End If
        
    Next
    
    Testa_Servico_Exclui = SUCESSO
    
    Exit Function
    
Erro_Testa_Servico_Exclui:
    
    Testa_Servico_Exclui = gErr
    
    Select Case gErr
        
        Case 98311
    
        Case 98322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SERVICO", gErr, objServicoSol.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
     
End Function

Function Testa_Servico_Inclui(ByVal ColServico As Collection, ByVal ColServicoSol As Collection, ByVal colservicoInclui As Collection) As Long
'Descobre os serviços da solicitação que podem ser incluídos

Dim bNaoAchou_BD As Boolean
Dim objServicoSol As New ClassServico
Dim objServico As New ClassServico

On Error GoTo Erro_Testa_Servico_Inclui
       
    'Descobrir os serviços que devem ser incluídos
    For Each objServico In ColServico
        For Each objServicoSol In ColServicoSol
            
            bNaoAchou_BD = True
            'Se existe no banco, sai.
            If objServicoSol.sProduto = objServico.sProduto Then
                bNaoAchou_BD = False
                Exit For
            End If
            
        Next
        'Se não existe --> Inclui
        If bNaoAchou_BD Then colservicoInclui.Add objServico
    Next
    
    Testa_Servico_Inclui = SUCESSO
    
    Exit Function
    
Erro_Testa_Servico_Inclui:
    
    Testa_Servico_Inclui = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
     
End Function

Function Testa_Existencia_Comprovante(ByVal ColServicoComp As Collection, ByVal objServicoSol As ClassServico, bNaoAchou_Comp As Boolean) As Long
'Verifica se o serviço da solicitação está associado a algum Comprovante

Dim objServicoComp As New ClassServico

On Error GoTo Erro_Testa_Existencia_Comprovante
    
    bNaoAchou_Comp = True
    
    For Each objServicoComp In ColServicoComp
    
        'Se o serviço do banco estiver sendo usado por algum comprovante
        If objServicoComp.sProduto = objServicoSol.sProduto Then
        
            bNaoAchou_Comp = False
            Exit For
            
        End If
        
    Next
    
    Testa_Existencia_Comprovante = SUCESSO
    
    Exit Function
    
Erro_Testa_Existencia_Comprovante:
    
    Testa_Existencia_Comprovante = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
    
End Function

Public Function Preenche_ColServicoSol(ByVal ColServicoSol As Collection, ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Preenche a coleção com os dados do serviço existente no banco

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim sProduto As String
Dim dQuant As Double
Dim objServicoSol As ClassServico

On Error GoTo Erro_Preenche_ColServicoSol
    
    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98309
    
    'Inicializa a variável
    sProduto = String(STRING_PRODUTO_CODIGO, 0)
    
    'Ler todos os servicos de SolServServico e colocar numa colecao
    lErro = Comando_ExecutarPos(lComando, "SELECT Produto, Quantidade FROM SolServServico WHERE NumIntDocOrigem=? ", 0, sProduto, dQuant, objSolicitacaoServico.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98195
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98196

    Do While lErro = AD_SQL_SUCESSO
        
        Set objServicoSol = New ClassServico
        
        objServicoSol.sProduto = sProduto
        objServicoSol.dQuant = dQuant
        
        'Preenche a coleção com o obj carregado
        ColServicoSol.Add objServicoSol
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98197

    Loop
        
    'Fecha o comando
    lComando = Comando_Fechar(lComando)
            
    Preenche_ColServicoSol = SUCESSO

    Exit Function

Erro_Preenche_ColServicoSol:

    Preenche_ColServicoSol = gErr

    Select Case gErr
        
        Case 98195, 98196, 98197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)
               
        Case 98309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
     
    'Fecha o comando
    lComando = Comando_Fechar(lComando)
        
    Exit Function

End Function
    
Public Function Preenche_ColServicoComp(ByVal ColServicoComp As Collection, ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Preenche a coleção com os dados do serviço existente no Comprovante associado a este serviço

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objServicoComp As ClassServico
Dim lCodigo As Long

On Error GoTo Erro_Preenche_ColServicoComp
    
    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98346
    
    'Inicializa a variável
    sProduto = String(STRING_PRODUTO_CODIGO, 0)
    
    'Ler todos os servicos de Comprovante associados a solicitacao e colocar numa colecao.
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Produto FROM CompServGR WHERE NumIntDocOrigem=? ", 0, lCodigo, sProduto, objSolicitacaoServico.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98198
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98199

    Do While lErro = AD_SQL_SUCESSO
        
        Set objServicoComp = New ClassServico
        
        objServicoComp.sProduto = sProduto
                        
        'Preenche a coleção com o obj carregado
        ColServicoComp.Add objServicoComp
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98200

    Loop
        
    'Fecha o comando
    lComando = Comando_Fechar(lComando)
            
    Preenche_ColServicoComp = SUCESSO

    Exit Function

Erro_Preenche_ColServicoComp:

    Preenche_ColServicoComp = gErr

    Select Case gErr
        
        Case 98198, 98199, 98200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
               
        Case 98346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
     
    'Fecha o comando
    lComando = Comando_Fechar(lComando)
        
    Exit Function

End Function
          
Public Function SolicitacaoServico_Lock(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Faz Lock de todas as tabelas necessárias para se gravar a Solicitação de Serviço

Dim lErro As Long
Dim alComando(1 To 9) As Long
Dim iIndice As Integer
Dim objSolServServico As ClassServico
Dim iCodigo As Integer
Dim lCodigo As Long
Dim iTipo As Integer
Dim sOrigemCidade As String
Dim sOrigemUF As String
Dim sDestinoCidade As String
Dim sDestinoUF As String
Dim lCliente As Long
Dim iOrigem As Integer
Dim iDestino As Integer
Dim dtData As Date

On Error GoTo Erro_SolicitacaoServico_Lock

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98262
    Next
        
    'Traz o Código do cliente, da origem e do destino
    'lock Tabela de Preço
    lErro = Comando_ExecutarPos(alComando(1), "SELECT DataVigencia, Cliente, Origem, Destino FROM TabPrecoGR WHERE Codigo = ? AND DataVigencia<=? ", 0, dtData, lCliente, iOrigem, iDestino, objSolicitacaoServico.lCodTabPreco, objSolicitacaoServico.dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then gError 98223

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98224
         
    'Se não existe dados associados com a Tabela de Preço  --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98225
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 98226
    
    'Joga o dado recolhido para objsolicitacaoservico
    objSolicitacaoServico.lCliente = lCliente
    
    'lock Cliente
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM Clientes WHERE Codigo = ? ", 0, lCodigo, lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 98227

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98228
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 98229
    
    'Inicializa as variáveis
    sOrigemCidade = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    sOrigemUF = String(STRING_ORIGEMDESTINO_UF, 0)
    
    'lock Origem
    lErro = Comando_ExecutarPos(alComando(3), "SELECT OrigemDestino, UF FROM OrigemDestino WHERE Codigo = ? ", 0, sOrigemCidade, sOrigemUF, iOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 98230

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98231
    
    lErro = Comando_LockShared(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 98192
    
    'Joga os dados recolhidos para objsolicitacaoservico
    objSolicitacaoServico.objEnderecoOrigem.sCidade = sOrigemCidade
    objSolicitacaoServico.objEnderecoOrigem.sSiglaEstado = sOrigemUF
    
    'Inicializa as variáveis
    sDestinoCidade = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    sDestinoUF = String(STRING_ORIGEMDESTINO_UF, 0)
    
    'lock Destino
    lErro = Comando_ExecutarPos(alComando(4), "SELECT OrigemDestino, UF FROM OrigemDestino WHERE Codigo = ? ", 0, sDestinoCidade, sDestinoUF, iDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 98189

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98190
    
    lErro = Comando_LockShared(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError 98191
    
    'Joga os dados recolhidos para objsolicitacaoservico
    objSolicitacaoServico.objEnderecoDestino.sCidade = sDestinoCidade
    objSolicitacaoServico.objEnderecoDestino.sSiglaEstado = sDestinoUF
    
    'Ler e lock Container
    lErro = Comando_ExecutarPos(alComando(5), "SELECT Tipo FROM TipoContainer WHERE Tipo=?", 0, iTipo, objSolicitacaoServico.iTipoContainer)
    If lErro <> AD_SQL_SUCESSO Then gError 98232

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98233
    
    lErro = Comando_LockShared(alComando(5))
    If lErro <> AD_SQL_SUCESSO Then gError 98234
    
    'Ler e lock TipoEmbalagem
    lErro = Comando_ExecutarPos(alComando(6), "SELECT Tipo FROM TipoEmbalagem WHERE Tipo=?", 0, iTipo, objSolicitacaoServico.iTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 98235

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98236
    
    lErro = Comando_LockShared(alComando(6))
    If lErro <> AD_SQL_SUCESSO Then gError 98237
    
    'Se o despachante estiver preenchido
    If objSolicitacaoServico.iDespachante <> 0 Then
    
        'Ler e lock Despachante
        lErro = Comando_ExecutarPos(alComando(7), "SELECT Codigo FROM Despachante WHERE Codigo=?", 0, iCodigo, objSolicitacaoServico.iDespachante)
        If lErro <> AD_SQL_SUCESSO Then gError 98238
    
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98239
        
        lErro = Comando_LockShared(alComando(7))
        If lErro <> AD_SQL_SUCESSO Then gError 98240
        
    End If
          
    'Ler e faz lock em todos os produtos do Grid de Serviços
    For Each objSolServServico In objSolicitacaoServico.ColServico
    
        'Verifica se o produto está associado a tabela de preço vigente
        lErro = Comando_ExecutarPos(alComando(8), "SELECT CodTabela FROM TabPrecoItensGR WHERE CodTabela=? AND DataVigencia=? AND Produto=?", 0, lCodigo, objSolicitacaoServico.lCodTabPreco, dtData, objSolServServico.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 98241
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98242
        
        'Se o produto não está associado a tabela de preço vigente
        If lErro = AD_SQL_SEM_DADOS Then gError 98188
        
        lErro = Comando_LockShared(alComando(8))
        If lErro <> AD_SQL_SUCESSO Then gError 98243
        
    Next
    
    'Se a programação de navio não estiver preenchida
    If objSolicitacaoServico.lCodProgNavio <> 0 Then
    
        'Ler e lock Programação de Navio
        lErro = Comando_ExecutarPos(alComando(9), "SELECT Codigo FROM ProgNavio WHERE Codigo=?", 0, lCodigo, objSolicitacaoServico.lCodProgNavio)
        If lErro <> AD_SQL_SUCESSO Then gError 98244
    
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98245
        
        lErro = Comando_LockShared(alComando(9))
        If lErro <> AD_SQL_SUCESSO Then gError 98246
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
        
    SolicitacaoServico_Lock = SUCESSO

    Exit Function

Erro_SolicitacaoServico_Lock:

    SolicitacaoServico_Lock = gErr

    Select Case gErr
    
        Case 98188
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ASSOCIADO_TABPRECO", gErr, objSolServServico.sProduto, objSolicitacaoServico.lCodTabPreco)
        
        Case 98189, 98190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, iDestino)
        
        Case 98191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORIGEMDESTINO", gErr, iDestino)
        
        Case 98192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORIGEMDESTINO", gErr, iOrigem)
        
        Case 98223, 98224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objSolicitacaoServico.lCodTabPreco)

        Case 98225
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_CADASTRADA1", gErr, objSolicitacaoServico.lCodTabPreco)
        
        Case 98226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABPRECO", gErr, objSolicitacaoServico.lCodTabPreco)
        
        Case 98227, 98228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", gErr, lCliente)
        
        Case 98229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", gErr, lCliente)
        
        Case 98230, 98231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, iOrigem)

        Case 98232, 98233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER", gErr, objSolicitacaoServico.iTipoContainer)

        Case 98234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOCONTAINER", gErr, objSolicitacaoServico.iTipoContainer)
        
        Case 98235, 98236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM", gErr, objSolicitacaoServico.iTipoEmbalagem)
        
        Case 98237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOEMBALAGEM", gErr, objSolicitacaoServico.iTipoEmbalagem)

        Case 98238, 98239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPACHANTE", gErr, objSolicitacaoServico.iDespachante)
        
        Case 98240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_DESPACHANTE", gErr, objSolicitacaoServico.iDespachante)

        Case 98241, 98242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objSolicitacaoServico.lCodTabPreco)
        
        Case 98243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABPRECOITENS", gErr, objSolicitacaoServico.lCodTabPreco)

        Case 98244, 98245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROGNAVIO", gErr, objSolicitacaoServico.lCodProgNavio)
        
        Case 98246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PROGNAVIO", gErr, objSolicitacaoServico.lCodProgNavio)
                   
        Case 98262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Verifica_Comprovante(objSolicitacaoServico As ClassSolicitacaoServico) As Long

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_Verifica_Comprovante
    
    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98325
        
    'Verificar se existem comprovantes de servico relacionados a esta solicitacao
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM CompServGR, SolicitacaoServico WHERE Numero=? AND SolicitacaoServico.FilialEmpresa=? AND CompServGR.NumIntDocOrigem=SolicitacaoServico.NumIntDoc", lCodigo, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98326
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98327
    
    'Se existe comprovante relacionado a solicitação
    If lErro = AD_SQL_SUCESSO Then
            
        'Avisa ao usuário os campos que poderam ser alterados
        vbMsgRet = Rotina_Aviso(vbOK, "AVISO_COMP_ASSOCIADO_SOLIC", objSolicitacaoServico.lNumero, lCodigo)
        
    End If
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Verifica_Comprovante = SUCESSO

    Exit Function

Erro_Verifica_Comprovante:

    Verifica_Comprovante = gErr

    Select Case gErr

        Case 98325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98326, 98327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Despachante_Le_NomeRed(objDespachante As ClassDespachante) As Long
'Lê os dados da tabela Despachante cujo nome reduzido foi passado.
'Devolve os dados em objDespachante.

Dim lErro As Long
Dim lComando As Long
Dim tDespachante As typeDespachante

On Error GoTo Erro_Despachante_Le_NomeRed

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98513
    
    'Inicializa
    tDespachante.sCGC = String(STRING_DESPACHANTE_CGC, 0)
    tDespachante.sNome = String(STRING_DESPACHANTE_NOME, 0)
    tDespachante.sNomeReduzido = String(STRING_DESPACHANTE_NOMEREDUZIDO, 0)

    'Lê os dados da tabela Despachante.
    lErro = Comando_Executar(lComando, "SELECT Codigo, CGC, Nome, NomeReduzido, Endereco FROM Despachante WHERE NomeReduzido = ? " _
    , tDespachante.iCodigo, tDespachante.sCGC, tDespachante.sNome, tDespachante.sNomeReduzido, tDespachante.lEndereco, objDespachante.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 98514

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98515
    
    ' Se não existe Despachante com o Código passado --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98516

    'Carrega os dados em objDespachante
    objDespachante.sCGC = tDespachante.sCGC
    objDespachante.sNome = tDespachante.sNome
    objDespachante.lEndereco = tDespachante.lEndereco
    objDespachante.iCodigo = tDespachante.iCodigo
        
    Call Comando_Fechar(lComando)

    Despachante_Le_NomeRed = SUCESSO

    Exit Function

Erro_Despachante_Le_NomeRed:

    Despachante_Le_NomeRed = gErr

    Select Case gErr

        Case 98513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98514, 98515
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPACHANTE2", gErr, objDespachante.sNomeReduzido)

        Case 98516
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabPrecoServico_Le1(ByVal objSolicitacaoServico As ClassSolicitacaoServico, ByVal sProduto As String) As Long
'Faz a verificação da validade do produto na Tabela de Preço

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_TabPrecoServico_Le1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98295
    
    'Lê os dados da tabela TabPrecoItensGR com o produto passado, verificando se combina
    'com a tabela de preço e data de vigencia em questão
    lErro = Comando_Executar(lComando, "SELECT CodTabela FROM TabPrecoGR, TabPrecoItensGR WHERE Codigo = ? And TabPrecoGR.DataVigencia <= ? AND Codtabela = Codigo AND TabPrecoGR.DataVigencia = TabPrecoItensGR.DataVigencia AND Produto = ? ORDER BY TabPrecoGr.DataVigencia DESC", lCodigo, objSolicitacaoServico.lCodTabPreco, objSolicitacaoServico.dtDataEmissao, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 98290
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98291
    
    'Se não encontrar --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98292
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    TabPrecoServico_Le1 = SUCESSO

    Exit Function

Erro_TabPrecoServico_Le1:

    TabPrecoServico_Le1 = gErr

    Select Case gErr
        
        Case 98290, 98291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objSolicitacaoServico.lCodTabPreco)
        
        Case 98292
            'Produto não relacionado com a tabela
            
        Case 98095
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function
   
End Function

Function TabPrecoServico_Le(ByVal objTabPreco As ClassTabPreco) As Long
'Le a tabela de Preços referentes a Solicitação de Serviço

Dim lErro As Long
Dim lCliente As Long
Dim iOrigem As Integer
Dim iDestino As Integer
Dim sClienteNomeRed As String
Dim sOrigemCidade As String
Dim sDestinoCidade As String
Dim sOrigemUF As String
Dim sDestinoUF As String
Dim lComando As Long

On Error GoTo Erro_TabPrecoServico_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98098
    
    'Inicializa as variáveis
    sClienteNomeRed = String(STRING_CLIENTES_NOMEREDUZIDO, 0)
    sOrigemCidade = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    sOrigemUF = String(STRING_ORIGEMDESTINO_UF, 0)
    sDestinoCidade = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    sDestinoUF = String(STRING_ORIGEMDESTINO_UF, 0)
    
    'Lê os dados da tabela TabPrecoGR, Clientes e OrigemDestino
    lErro = Comando_Executar(lComando, "SELECT Cliente, NomeReduzido, Origem.OrigemDestino, Origem.UF, Destino.OrigemDestino, Destino.UF FROM TabPrecoGR, Clientes, OrigemDestino as Origem, OrigemDestino as Destino WHERE TabPrecoGR.Codigo = ? AND DataVigencia<=? AND TabPrecoGR.Cliente = Clientes.Codigo AND TabPrecoGR.Origem = Origem.Codigo AND TabPrecoGR.Destino = Destino.Codigo", _
    lCliente, sClienteNomeRed, sOrigemCidade, sOrigemUF, sDestinoCidade, sDestinoUF, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 98099

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98100
         
    'Se não existe dados associados com a Tabela de Preço  --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98101
    
    'Carrega o obj com os dados recolhidos no banco
    objTabPreco.lCliente = lCliente
    objTabPreco.sClienteNomeRed = sClienteNomeRed
    objTabPreco.sDestinoCidade = sDestinoCidade
    objTabPreco.sDestinoUF = sDestinoUF
    objTabPreco.sOrigemCidade = sOrigemCidade
    objTabPreco.sOrigemUF = sOrigemUF
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    TabPrecoServico_Le = SUCESSO

    Exit Function

Erro_TabPrecoServico_Le:

    TabPrecoServico_Le = gErr

    Select Case gErr

        Case 98098
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98099, 98100
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)
        
        Case 98101
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function
   
End Function

Function SolServServico_Le(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long 'Lê todos os serviços associados a Solicitação de Serviço cujo código foi passado

Dim lErro As Long
Dim lComando As Long
Dim objSolServServico As ClassServico
Dim sProduto As String
Dim dQuant As Double
Dim sDescricao As String

On Error GoTo Erro_SolServServico_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98091
    
    'Inicializa as variáveis
    sProduto = String(STRING_SOLSERVSERVICO_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTOS_DESCRICAO, 0)
    
    'Lê os dados da tabela SolServServico e Produtos com o NumIntDoc da solicitação
    lErro = Comando_Executar(lComando, "SELECT Produto, Quantidade, Descricao FROM SolServServico, Produtos WHERE NumIntDocOrigem = ? AND SolServServico.Produto = Produtos.Codigo", sProduto, dQuant, sDescricao, objSolicitacaoServico.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98092

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98093
         
    'Se não existe Serviço associado com a Solicitação de Serviço  --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98095
         
    'Enquanto achar serviço
    Do While lErro = AD_SQL_SUCESSO
    
        'Inicializa SolServServico
        Set objSolServServico = New ClassServico
        
        'Carrega os dados em objSolServServico
        objSolServServico.sProduto = sProduto
        objSolServServico.sDescricao = sDescricao
        objSolServServico.dQuant = dQuant
        
        'Preenche a coleção com o obj carregado
        objSolicitacaoServico.ColServico.Add objSolServServico
    
        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98094
        
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    SolServServico_Le = SUCESSO

    Exit Function

Erro_SolServServico_Le:

    SolServServico_Le = gErr

    Select Case gErr

        Case 98091
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98092, 98093, 98094
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)
        
        Case 98095
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SolicitacaoServico_Le(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Le os dados da tabela Solicitação de Serviço

Dim lErro As Long
Dim TypeSolServ As typeSolicitacaoServico
Dim sDespachante As String
Dim sCliente As String
Dim lComando As Long

On Error GoTo Erro_SolicitacaoServico_Le
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98090
    
    With TypeSolServ
        
        'Inicializar as variáveis
        .sNumReferencia = String(STRING_SOLICITACAOSERVICO_NUMREFERENCIA, 0)
        .sMaterial = String(STRING_SOLICITACAOSERVICO_MATERIAL, 0)
        .sUM = String(STRING_SOLICITACAOSERVICO_UM, 0)
        .sBooking = String(STRING_SOLICITACAOSERVICO_BOOKING, 0)
        .sObservacao = String(STRING_SOLICITACAOSERVICO_OBSERVACAO, 0)
        sDespachante = String(STRING_DESPACHANTE_NOMEREDUZIDO, 0)
        sCliente = String(STRING_CLIENTES_NOMEREDUZIDO, 0)
        .sPorto = String(STRING_SOLICITACAO_PORTO, 0)
        
        'Le a View SolicitacaoRel para recolher do Banco todos os dados da tabela SolicitacaoServico, o nome reduzido do Cliente e do Despachante.
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, CodTabPreco, Cliente, DataEmissao, NumReferencia, DataPedido, TipoOperacao, Despachante, Material, QuantMaterial, UM, ValorMercadoria, TipoEmbalagem, TipoContainer, CodProgNavio, Booking, DataPrevInicio, HoraPrevInicio, DataPrevFim, HoraPrevFim, Observacao, EnderecoOrigem, EnderecoDestino, Porto " _
        & ",CliNomeRed , DesNomeRed FROM SolicitacaoRel Where Numero = ? AND FilialEmpresa = ?", .lNumIntDoc, .lCodTabPreco, .lCliente, .dtDataEmissao, .sNumReferencia, .dtDataPedido, .iTipoOperacao, .iDespachante, .sMaterial, .dQuantMaterial, .sUM, .dValorMercadoria, .iTipoEmbalagem, .iTipoContainer, .lCodProgNavio, .sBooking, .dtDataPrevInicio, .dHoraPrevInicio, _
        .dtDataPrevFim, .dHoraPrevFim, .sObservacao, .lEnderecoOrigem, .lEnderecoDestino, .sPorto, sCliente, sDespachante, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 98075
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98076
        
        'Se não encontrou --> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 98085
        
        'Inicializa os objEndereco
        Set objSolicitacaoServico.objEnderecoOrigem = New ClassEndereco
        Set objSolicitacaoServico.objEnderecoDestino = New ClassEndereco
        
        'Carrega o obj com os endereços recolhidos no banco
        objSolicitacaoServico.objEnderecoOrigem.lCodigo = .lEnderecoOrigem
        objSolicitacaoServico.objEnderecoDestino.lCodigo = .lEnderecoDestino
        
        'Lê o endereço com o código passado
        lErro = CF("Endereco_Le", objSolicitacaoServico.objEnderecoOrigem)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 98280
        
        'Se não encontrou o endereço --> erro
        If lErro = 12309 Then gError 98281
        
        'Lê o endereço com o código passado
        lErro = CF("Endereco_Le", objSolicitacaoServico.objEnderecoDestino)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 98282
        
        'Se não encontrou o endereço --> erro
        If lErro = 12309 Then gError 98283
        
        'Carrega o objSolicitacaoServico
        objSolicitacaoServico.dQuantMaterial = .dQuantMaterial
        objSolicitacaoServico.dtDataEmissao = .dtDataEmissao
        objSolicitacaoServico.dtDataPedido = .dtDataPedido
        objSolicitacaoServico.dtDataPrevFim = .dtDataPrevFim
        objSolicitacaoServico.dtDataPrevInicio = .dtDataPrevInicio
        objSolicitacaoServico.dtHoraPrevFim = .dHoraPrevFim
        objSolicitacaoServico.dtHoraPrevInicio = .dHoraPrevInicio
        objSolicitacaoServico.dValorMercadoria = .dValorMercadoria
        objSolicitacaoServico.iTipoContainer = .iTipoContainer
        objSolicitacaoServico.iTipoEmbalagem = .iTipoEmbalagem
        objSolicitacaoServico.iTipoOperacao = .iTipoOperacao
        objSolicitacaoServico.lCodProgNavio = .lCodProgNavio
        objSolicitacaoServico.lCodTabPreco = .lCodTabPreco
        objSolicitacaoServico.lNumIntDoc = .lNumIntDoc
        objSolicitacaoServico.sBooking = .sBooking
        objSolicitacaoServico.sMaterial = .sMaterial
        objSolicitacaoServico.sNumReferencia = .sNumReferencia
        objSolicitacaoServico.sObservacao = .sObservacao
        objSolicitacaoServico.sUM = .sUM
        objSolicitacaoServico.sPorto = .sPorto
        
    End With
    
    'Carrega o obj com os outros campos que não fazem parte da tabela solicitação
    objSolicitacaoServico.sClienteNomeRed = sCliente
    objSolicitacaoServico.sDespachanteNomeRed = sDespachante
   
    Call Comando_Fechar(lComando)
    
    SolicitacaoServico_Le = SUCESSO

    Exit Function

Erro_SolicitacaoServico_Le:

    SolicitacaoServico_Le = gErr

    Select Case gErr

        Case 98075, 98076
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
            
        Case 98085, 98280, 98282
        
        Case 98090
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 98281, 98283
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function CompServGR_Le(ByVal objCompServ As ClassCompServ) As Long
'Lê o Comprovante Serv. a partir do
'Código e FilialEmpresa passados como parametros

Dim lErro As Long, lComando As Long
Dim tCompServ As typeCompServ
Dim lCliente As Long
Dim iTipoOperacao As Integer

On Error GoTo Erro_CompServGR_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97416

    'inicializando
    tCompServ.sCodigoContainer = String(STRING_CODCONTAINER, 0)
    tCompServ.sLacre = String(STRING_LACRE, 0)
    tCompServ.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServ.sProduto = String(STRING_PRODUTO_CODIGO, 0)
    tCompServ.sUM = String(STRING_UM_NOME, 0)

    'Busca na tabela CompServGR a Cotacao com Código e FilialEmpresa informados
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT CompServGR.NumIntDoc, CompServGR.NumIntDocOrigem, CompServGR.FilialEmpresa, Codigo, CompServGR.DataEmissao, Produto, CompServGR.QuantMaterial, CompServGR.UM, CompServGR.ValorMercadoria," _
    & "FretePeso, Pedagio, AdValoren, DataDemurrage, CodigoContainer, ValorContainer, Tara, Lacre, CompServGR.Observacao, NumIntNota, Situacao, Quantidade, Cliente,TipoOperacao " _
    & "FROM CompServGR, SolicitacaoServico " _
    & "WHERE Codigo=? AND CompServGR.FilialEmpresa=? AND SolicitacaoServico.NumIntDoc = CompServGR.NumIntDocOrigem", tCompServ.lNumIntDoc, tCompServ.lNumIntDocOrigem, tCompServ.iFilialEmpresa, tCompServ.lCodigo, tCompServ.dtDataEmissao, _
    tCompServ.sProduto, tCompServ.dQuantMaterial, tCompServ.sUM, tCompServ.dValorMercadoria, tCompServ.dFretePeso, tCompServ.dPedagio, _
    tCompServ.dAdValoren, tCompServ.dtDataDemurrage, tCompServ.sCodigoContainer, tCompServ.dValorContainer, tCompServ.dTara, tCompServ.sLacre, tCompServ.sObservacao, tCompServ.lNumIntNota, tCompServ.iSituacao, tCompServ.dQuantidade, lCliente, iTipoOperacao, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 97417

    'Se posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97418

    'Se não encontrou o CompServ ==> erro 97419
    If lErro = AD_SQL_SEM_DADOS Then gError 97419

    'Carrega o obj com o conteudo lido do bd
    With objCompServ

        .dAdValoren = tCompServ.dAdValoren
        .dFretePeso = tCompServ.dFretePeso
        .dPedagio = tCompServ.dPedagio
        .dQuantMaterial = tCompServ.dQuantMaterial
        .dTara = tCompServ.dTara
        .dtDataDemurrage = tCompServ.dtDataDemurrage
        .dtDataEmissao = tCompServ.dtDataEmissao
        .dValorMercadoria = tCompServ.dValorMercadoria
        .lNumIntDoc = tCompServ.lNumIntDoc
        .lNumIntDocOrigem = tCompServ.lNumIntDocOrigem
        .sCodigoContainer = tCompServ.sCodigoContainer
        .dValorContainer = tCompServ.dValorContainer
        .sLacre = tCompServ.sLacre
        .sObservacao = tCompServ.sObservacao
        .sProduto = tCompServ.sProduto
        .sUM = tCompServ.sUM
        .iSituacao = tCompServ.iSituacao
        .lNumIntNota = tCompServ.lNumIntNota
        .dQuantidade = tCompServ.dQuantidade
        .lCliente = lCliente
        .iTipoOperacao = iTipoOperacao

    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServGR_Le = SUCESSO

    Exit Function

Erro_CompServGR_Le:

    CompServGR_Le = gErr

    Select Case gErr

        Case 97416
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97417, 97418
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV2", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 97419
            'NAO ACHOU COMPSERV !!

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CompServNFiscal_Le(ByVal objCompServ As ClassCompServ) As Long
'Lê o Comprovante Serv numa view para a nota fiscal
'Código e FilialEmpresa passados como parametros

Dim lErro As Long, lComando As Long
Dim tCompServ As typeCompServ
Dim lCliente As Long

On Error GoTo Erro_CompServNFiscal_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99068

    'inicializando
    tCompServ.sCodigoContainer = String(STRING_CODCONTAINER, 0)
    tCompServ.sLacre = String(STRING_LACRE, 0)
    tCompServ.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServ.sProduto = String(STRING_PRODUTO_CODIGO, 0)
    tCompServ.sUM = String(STRING_UM_NOME, 0)

    'Busca na view ComprovServNFiscal a Cotacao com Código e FilialEmpresa informados
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, FilialEmpresa, Codigo, DataEmissao, Produto, QuantMaterial, UM, ValorMercadoria," _
    & "FretePeso, Pedagio, AdValoren, DataDemurrage, CodigoContainer, ValorContainer, Tara, Lacre, Observacao, NumIntNota, Situacao, Quantidade, ClienteSol " _
    & "FROM ComprovNFiscal " _
    & "WHERE Codigo=? AND FilialEmpresa=?", tCompServ.lNumIntDoc, tCompServ.lNumIntDocOrigem, tCompServ.iFilialEmpresa, tCompServ.lCodigo, tCompServ.dtDataEmissao, _
    tCompServ.sProduto, tCompServ.dQuantMaterial, tCompServ.sUM, tCompServ.dValorMercadoria, tCompServ.dFretePeso, tCompServ.dPedagio, _
    tCompServ.dAdValoren, tCompServ.dtDataDemurrage, tCompServ.sCodigoContainer, tCompServ.dValorContainer, tCompServ.dTara, tCompServ.sLacre, tCompServ.sObservacao, tCompServ.lNumIntNota, tCompServ.iSituacao, tCompServ.dQuantidade, lCliente, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99069

    'Se posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99070

    'Se não encontrou o CompServ
    If lErro = AD_SQL_SEM_DADOS Then gError 99071

    'Carrega o obj com o conteudo lido do bd
    With objCompServ

        .dAdValoren = tCompServ.dAdValoren
        .dFretePeso = tCompServ.dFretePeso
        .dPedagio = tCompServ.dPedagio
        .dQuantMaterial = tCompServ.dQuantMaterial
        .dTara = tCompServ.dTara
        .dtDataDemurrage = tCompServ.dtDataDemurrage
        .dtDataEmissao = tCompServ.dtDataEmissao
        .dValorMercadoria = tCompServ.dValorMercadoria
        .lNumIntDoc = tCompServ.lNumIntDoc
        .lNumIntDocOrigem = tCompServ.lNumIntDocOrigem
        .sCodigoContainer = tCompServ.sCodigoContainer
        .dValorContainer = tCompServ.dValorContainer
        .sLacre = tCompServ.sLacre
        .sObservacao = tCompServ.sObservacao
        .sProduto = tCompServ.sProduto
        .sUM = tCompServ.sUM
        .iSituacao = tCompServ.iSituacao
        .lNumIntNota = tCompServ.lNumIntNota
        .dQuantidade = tCompServ.dQuantidade
        .lCliente = lCliente
        
    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServNFiscal_Le = SUCESSO

    Exit Function

Erro_CompServNFiscal_Le:

    CompServNFiscal_Le = gErr

    Select Case gErr

        Case 99068
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99069, 99070
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV2", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 99071
            'NAO ACHOU COMPSERV !!

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Documento_Le_NomeTsk(sNomeTsk As String, ByVal sNome As String) As Long
'Le o item e retorna o nome do documento associado a ele no sNomeTsk

Dim lErro As Long, lComando As Long

On Error GoTo Erro_Documento_Le_NomeTsk

    sNomeTsk = String(STRING_DOCUMENTO_DOCUMENTO, 0)

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98561
    
    'busca o item
    lErro = Comando_Executar(lComando, "SELECT Documento from Documento where NomeReduzido = ?", sNomeTsk, sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 98562
        
    'posiciona no dito cujo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98563
    
    'se nao achou o item --> erro, possível
    'problema de integridade
    If lErro = AD_SQL_SEM_DADOS Then gError 98564
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Documento_Le_NomeTsk = SUCESSO
    
    Exit Function

Erro_Documento_Le_NomeTsk:

    Documento_Le_NomeTsk = gErr
    
    Select Case gErr
        
        Case 98561
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98562, 98563
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO1", gErr)
        
        Case 98564
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_EXISTENTE", gErr, sNome)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function SolServServTabPreItens_Le(ByVal objSolServ As ClassSolicitacaoServico, ByVal objProduto As ClassProduto, dPrecoProd As Double, dQtdProd As Double) As Long
'Funcao que le a quantidade e o preco de um produto (objproduto), relacionado com um solicitacao de servico (objsolserv)

Dim lErro As Long, lComando As Long, lTab As Long
Dim dPreco As Double, dQtd As Double
Dim dtVig As Date, dtEmiSol As Date

On Error GoTo Erro_SolServServTabPreItens_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98617

    'pega a data da solicitacao e o codigo da tabela de preco
    lErro = Comando_Executar(lComando, "SELECT DataEmissao, CodTabPreco FROM solicitacaoservico WHERE Numero = ? and FilialEmpresa = ?", dtEmiSol, lTab, objSolServ.lNumero, objSolServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98689
    
    'colocar esse erro la embaixo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98702
    
    If lErro <> AD_SQL_SUCESSO Then gError 98690
    
    'pega a data de vigencia
    lErro = Comando_Executar(lComando, "SELECT z.datavigencia FROM tabprecoitensgr z WHERE z.codtabela = ? AND z.produto = ? AND z.datavigencia <= ? ORDER BY z.DataVigencia DESC", dtVig, lTab, objProduto.sCodigo, dtEmiSol)
    If lErro <> AD_SQL_SUCESSO Then gError 98691

    'colocar esse erro la embaixo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98703
    
    If lErro <> AD_SQL_SUCESSO Then gError 98692

    'captura finalmente a quantidade e o preco
    lErro = Comando_Executar(lComando, "SELECT a.preco, b.quantidade FROM tabprecoitensgr a, solservservico b, solicitacaoservico c WHERE a.produto = ? AND b.produto = a.produto AND c.filialempresa = ? AND c.numero = ? AND c.numintdoc = b.numintdocorigem AND a.datavigencia = ? AND a.CodTabela = ? ORDER BY a.Codtabela Desc", _
           dPreco, dQtd, objProduto.sCodigo, objSolServ.iFilialEmpresa, objSolServ.lNumero, dtVig, lTab)
        
    If lErro <> AD_SQL_SUCESSO Then gError 98618
        
    'posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98619
        
    'Se nao achou...
    If lErro <> AD_SQL_SUCESSO Then gError 98620
    
    'Atribui a quantidade e o preco nas variaveis passadas por referencia
    dPrecoProd = dPreco
    dQtdProd = dQtd
                  
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    SolServServTabPreItens_Le = SUCESSO
    
    Exit Function
    
Erro_SolServServTabPreItens_Le:

    SolServServTabPreItens_Le = gErr
    
    Select Case gErr
        
        Case 98617
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error$)
        
        Case 98618, 98619
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICICITACAOSERVICO1", gErr)
        
        Case 98620
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERVSERV_TABPREITENS_NAO_ENCONTRADO", gErr, objProduto.sCodigo, objSolServ.lNumero, objSolServ.iFilialEmpresa, dtVig)
        
        Case 98689
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO3", gErr, objSolServ.lNumero, objSolServ.iFilialEmpresa)
        
        'esse erro nao ocorrera no caso dessa tela
        Case 98690
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolServ.lNumero, objSolServ.iFilialEmpresa)
        
        Case 98691
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr)
        
        Case 98692
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_FOI_POSSIVEL_OBTER_DTVIGENCIA", gErr, lTab, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Documento_Le_NomeRed(objDocumento As ClassDocumento) As Long
'Lê os dados da tabela Documento cujo Nome Reduzido foi passado.

Dim lErro As Long, lComando As Long
Dim tDocumento As typeDocumento

On Error GoTo Erro_Documento_Le_NomeRed

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98545
     
    With tDocumento
    
        'Inicializa
        .sDescricao = String(STRING_DOCUMENTO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_DOCUMENTO_NOMEREDUZIDO, 0)
        .sDocumento = String(STRING_DOCUMENTO_DOCUMENTO, 0)
        'Lê os dados do Documento.
        lErro = Comando_Executar(lComando, "SELECT Descricao, Codigo, TipoDoc,Documento FROM Documento WHERE NomeReduzido = ? ", .sDescricao, .iCodigo, .iTipoDoc, .sDocumento, objDocumento.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 98546
    
        'posiciona no registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98547
    
        'Se não existe Documento com o NomeRed passado --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 98548
    
        'Carrega os dados em objDocumento
        objDocumento.iTipoDoc = .iTipoDoc
        objDocumento.sDescricao = .sDescricao
        objDocumento.iCodigo = .iCodigo
        objDocumento.sDocumento = .sDocumento
                
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Documento_Le_NomeRed = SUCESSO

    Exit Function

Erro_Documento_Le_NomeRed:

    Documento_Le_NomeRed = gErr

    Select Case gErr

        Case 98545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98546, 98547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO2", gErr, objDocumento.sNomeReduzido)

        Case 98548
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TP_Doc_Le(ByVal objDoc As ClassDocumento, ByVal sDoc As String) As Long
'Lê o tipo Doc com Código ou NomeRed

Dim eTipoDoc As enumTipo
Dim lErro As Long

On Error GoTo Erro_TP_Doc_Le

    'Verifica se o parametro passado eh numerico
    'ou string...
    eTipoDoc = Tipo_Doc(sDoc)

    Select Case eTipoDoc

        'Se for string, entao eh nomereduzido..
        Case TIPO_STRING
            
            objDoc.sNomeReduzido = sDoc
            
            'Le documento a partir do nomered
            lErro = Documento_Le_NomeRed(objDoc)
            If lErro <> SUCESSO And lErro <> 98548 Then gError 98549
            
            If lErro <> SUCESSO Then gError 98550
                          
        'Se for numerico, entao eh codigo
        Case TIPO_CODIGO
    
            objDoc.iCodigo = StrParaInt(sDoc)
            
            'Le documento a partir do codigo
            lErro = CF("Documento_Le", objDoc)
            If lErro <> SUCESSO And lErro <> 98026 Then gError 98551
            
            If lErro <> SUCESSO Then gError 98552
            
        'Caso seja algo diferente dos anteriores,
        'erro...
        Case TIPO_OVERFLOW
    
            gError 98553
    
        Case TIPO_DECIMAL
    
            gError 98554
    
        Case TIPO_NAO_POSITIVO
    
            gError 98555
        
        Case TIPO_VAZIO
            
            gError 98180
            
    End Select

    TP_Doc_Le = SUCESSO

    Exit Function

Erro_TP_Doc_Le:

    TP_Doc_Le = gErr

    Select Case gErr
        
        Case 98549, 98551, 98550, 98552, 98180
        
        Case 98553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, sDoc)

        Case 98554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sDoc)

        Case 98555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Public Function Solicitacao_Servico_Le_NumIntDoc(ByVal objSolServ As ClassSolicitacaoServico) As Long
'Le a solicitacao de Servico atraves do numintdoc passado
'no obj... visa retornar o codigo para visualizacao
'na tela.

Dim lErro As Long, lComando As Long, lCodigo As Long
Dim iFilial As Integer

On Error GoTo Erro_Solicitacao_Servico_Le_NumIntDoc

    'abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97469

    'busca na tabela a solserv com o numintdoc passado
    lErro = Comando_Executar(lComando, "SELECT Numero, FilialEmpresa FROM SolicitacaoServico where NumIntDoc = ?", lCodigo, iFilial, objSolServ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97470
    
    'posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97471
    
    If lErro <> AD_SQL_SUCESSO Then gError 97472
    
    'Carrega no obj os dados recolhidos da tabela
    objSolServ.lNumero = lCodigo
    objSolServ.iFilialEmpresa = iFilial
            
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Solicitacao_Servico_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Solicitacao_Servico_Le_NumIntDoc:

    Solicitacao_Servico_Le_NumIntDoc = gErr
    
    Select Case gErr
    
        Case 97469
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 97470, 97471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO", gErr, objSolServ.lNumero)
    
        Case 97472
            'NAO ACHOU!!
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CompServGR_Le_CompServItem(ByVal objCompServ As ClassCompServ) As Long
'Le os CompServItens relacionados com o comprovante em questao
'armazenando-os na colecao

Dim lErro As Long, lComando As Long
Dim sCampos As String, sDoc As String
Dim tCompServItem As typeCompServItem
Dim objCompServItem As ClassCompServItem
Dim objDoc As New ClassDocumento

On Error GoTo Erro_CompServGR_Le_CompServItem

    'Abrindo Comando..
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98596

    'Inicializando as strings do type
    tCompServItem.sDocExtNumero = String(STRING_DOC_NUMERO, 0)
    tCompServItem.sDocIntNumero = String(STRING_DOC_NUMERO, 0)
    tCompServItem.sMotorista = String(STRING_MOTORISTA, 0)
    tCompServItem.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServItem.sPlacaCaminhao = String(STRING_PLACA, 0)
    tCompServItem.sPlacaCarreta = String(STRING_PLACA, 0)
         
    'Para auxiliar o select... guardo os campos em uma string
    sCampos = "NumIntDoc,NumIntDocOrigem,CodItemServico,DataPrev, " _
    & "HoraPrev,DataInicio,HoraInicio,DataFim,HoraFim,DocIntTipo,DocIntNumero, " _
    & "DocIntDataEmissao,DocExtTipo,DocExtNumero,DocExtDataEmissao,DocExtDataRec, " _
    & "DocExtHoraRec,PlacaCaminhao,PlacaCarreta,Motorista,Observacao"
    
    With tCompServItem
    
        'obtem um conjunto de registros da tabela compservitemgr
        'relacionados com o compserv
        lErro = Comando_Executar(lComando, "SELECT " & sCampos & " FROM CompServItemGR WHERE NumIntDocOrigem = ? ORDER BY NumIntDoc", _
        .lNumIntDoc, .lNumIntDocOrigem, .iCodItemServico, .dtDataPrev, .dHoraPrev, .dtDataInicio, .dHoraInicio, .dtDataFim, .dHoraFim, _
        .iDocIntTipo, .sDocIntNumero, .dtDocIntDataEmissao, .iDocExtTipo, .sDocExtNumero, .dtDocExtDataEmissao, .dtDocExtDataRec, .dDocExtHoraRec, .sPlacaCaminhao, .sPlacaCarreta, .sMotorista, .sObservacao, objCompServ.lNumIntDoc)
       
    End With
        
    If lErro <> SUCESSO Then gError 98597
        
    'posiciona-se no primeiro registro do conjunto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98598

    'preenche a colecao com os dados antes lidos
    'um obj eh utilizado para que o armazenamento
    'seja feito
    Do While lErro = AD_SQL_SUCESSO
        
        'instancia o objcompservitem
        Set objCompServItem = New ClassCompServItem
        
        With objCompServItem
            
            'carrega os dados no obj instanciado anteriormente..
            .dtDataFim = tCompServItem.dtDataFim
            .dtDataInicio = tCompServItem.dtDataInicio
            .dtDataPrev = tCompServItem.dtDataPrev
            .dtDocExtDataEmissao = tCompServItem.dtDocExtDataEmissao
            .dtDocExtDataRec = tCompServItem.dtDocExtDataRec
            .dtDocExtHoraRec = tCompServItem.dDocExtHoraRec
            .dtDocIntDataEmissao = tCompServItem.dtDocIntDataEmissao
            .dtHoraFim = tCompServItem.dHoraFim
            .dtHoraInicio = tCompServItem.dHoraInicio
            .dtHoraPrev = tCompServItem.dHoraPrev
            .iCodItemServico = tCompServItem.iCodItemServico
            .lNumIntDoc = tCompServItem.lNumIntDoc
            .lNumIntDocOrigem = tCompServItem.lNumIntDocOrigem
            .iDocExtTipo = tCompServItem.iDocExtTipo
            .iDocIntTipo = tCompServItem.iDocIntTipo
            .sDocExtNumero = tCompServItem.sDocExtNumero
            .sDocIntNumero = tCompServItem.sDocIntNumero
            
            If tCompServItem.iDocExtTipo <> 0 Then
            
                objDoc.iCodigo = tCompServItem.iDocExtTipo
                            
                'le o documento externo para obter o nomereduzido
                lErro = CF("Documento_Le", objDoc)
                If lErro <> SUCESSO Then gError 98599
                            
                .sDocExtTipo = objDoc.sNomeReduzido
                                  
            End If
            
            If tCompServItem.iDocIntTipo <> 0 Then
            
                objDoc.iCodigo = tCompServItem.iDocIntTipo
                
                'le o documento interno para obter o nomereduzido
                lErro = CF("Documento_Le", objDoc)
                If lErro <> SUCESSO Then gError 98600
                            
                .sDocIntTipo = objDoc.sNomeReduzido
            
            End If
            
            .sMotorista = tCompServItem.sMotorista
            .sObservacao = tCompServItem.sObservacao
            .sPlacaCaminhao = tCompServItem.sPlacaCaminhao
            .sPlacaCarreta = tCompServItem.sPlacaCarreta
              
        End With
        
        'adiciona o obj carregado na colecao...
        objCompServ.colCompServItem.Add objCompServItem

        'busca o proximo registro do conjunto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98601
    
    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServGR_Le_CompServItem = SUCESSO

    Exit Function

Erro_CompServGR_Le_CompServItem:

    CompServGR_Le_CompServItem = gErr

    Select Case gErr
        
        Case 98596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98597, 98598, 98601
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERVITEM", gErr)
        
        Case 98599, 98600
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function TabPreco_Le2(ByVal objTabPreco As ClassTabPreco, ByVal objSolServ As ClassSolicitacaoServico) As Long
'Lê os dados da tabela TabPreco cujo código
'foi passado.
'A data de vigencia eh a maior menor do que a data de emissao da solicitacao

Dim lErro As Long
Dim lComando As Long
Dim tTabPreco As typeTabPreco
Dim dtDataVig As Date

On Error GoTo Erro_TabPreco_Le2

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98695
     
    With tTabPreco
    
        'Inicializa
        .sObservacao = String(STRING_TABPRECO_OBSERVACAO, 0)
    
        'pega a data de vigencia
        lErro = Comando_Executar(lComando, "SELECT MAX(z.datavigencia) FROM tabprecogr z WHERE z.codigo = ? AND z.datavigencia <= ?", dtDataVig, objSolServ.lCodTabPreco, objSolServ.dtDataEmissao)
        If lErro <> AD_SQL_SUCESSO Then gError 98696

        lErro = Comando_BuscarPrimeiro(lComando)
        'colocar esse erro la embaixo
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98701
        
        If lErro <> AD_SQL_SUCESSO Then gError 98697
            
        'Lê os dados da tabela TabPreco.
        lErro = Comando_Executar(lComando, "SELECT AdValoren, Pedagio, Destino, Origem, Cliente, Observacao FROM TabPrecoGR WHERE Codigo = ? and DataVigencia = ?" _
        , .dAdValoren, .dPedagio, .iDestino, .iOrigem, .lCliente, .sObservacao, objSolServ.lCodTabPreco, dtDataVig)
        If lErro <> AD_SQL_SUCESSO Then gError 98698
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98699
    
        'Se não existe Tabela de Preço com o Código e a data passada --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 98700
    
        'Carrega os dados em objTabPreco
        objTabPreco.lCliente = .lCliente
        objTabPreco.dAdValoren = .dAdValoren
        objTabPreco.dPedagio = .dPedagio
        objTabPreco.iDestino = .iDestino
        objTabPreco.iOrigem = .iOrigem
        objTabPreco.sObservacao = .sObservacao
    
    End With
    
    Call Comando_Fechar(lComando)

    TabPreco_Le2 = SUCESSO

    Exit Function

Erro_TabPreco_Le2:

    TabPreco_Le2 = gErr

    Select Case gErr

        Case 98695
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98696, 98698, 98699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objSolServ.lCodTabPreco)

        Case 98687
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_FOI_POSSIVEL_OBTER_DTVIGENCIA2", gErr, objSolServ.lCodTabPreco)
        
        Case 98700
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Tipo_Doc(ByVal sText As String) As enumTipo
'Funcao que verifica o tipo do parametro sText

    If Len(Trim(sText)) = 0 Then
        Tipo_Doc = TIPO_VAZIO
    ElseIf Not IsNumeric(sText) Or Len(Trim(sText)) > 5 Then
        Tipo_Doc = TIPO_STRING
    ElseIf Int(CDbl(sText)) <> CDbl(sText) Then
        Tipo_Doc = TIPO_DECIMAL
    ElseIf CDbl(sText) <= 0 Then
        Tipo_Doc = TIPO_NAO_POSITIVO
    Else
        Tipo_Doc = TIPO_CODIGO
    End If

End Function

Public Function SolServ_Lock(ByVal objSolServ As ClassSolicitacaoServico) As Long
'Rotina executada dentro de uma TRANSACAO!!!!!!!!!!!!
'Faz lock shared da Solicitacao de Servico passada como parametro

Dim lErro As Long, lComando As Long, lNumIntDocDummy As Long, lNumero As Long
Dim iFilial As Integer

On Error GoTo Erro_SolServ_Lock
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98722
    
    'Pesquisa o Solicitacao de Servico no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT NumIntDoc, Numero, FilialEmpresa FROM SolicitacaoServico WHERE NumIntDoc = ?", lNumIntDocDummy, lNumero, iFilial, objSolServ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98643

    'Se posiciona no registro que vai receber o lock
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98644

    'Se não encontrou a Solicitacao --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98645
    
    'Locka efetivamente a solicitacao
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 98646
    
    Call Comando_Fechar(lComando)
    
    objSolServ.lNumero = lNumero
    objSolServ.iFilialEmpresa = iFilial
    
    SolServ_Lock = SUCESSO

    Exit Function

Erro_SolServ_Lock:

    SolServ_Lock = gErr

    Select Case gErr

        Case 98643, 98644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)

        Case 98645
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_CADASTRADA", gErr)

        Case 98646
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SOLICITACAOSERVICO", gErr)
        
        Case 98722
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Documento_Lock(ByVal objDoc As ClassDocumento) As Long
'Rotina executada dentro de uma TRANSACAO!!!!!!!!!!!!
'Faz lock shared do Documento

Dim lErro As Long, lComando As Long
Dim iCodDummy As Integer

On Error GoTo Erro_Documento_Lock
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98725
    
    'Pesquisa o Documento no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Documento WHERE NomeReduzido = ?", iCodDummy, objDoc.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 98726

    'Se posiciona no registro que vai receber o lock
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98727

    'Se não encontrou o Documento --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98728
    
    'Locka efetivamente o Documento
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 98729
    
    Call Comando_Fechar(lComando)
    
    objDoc.iCodigo = iCodDummy
    
    Documento_Lock = SUCESSO

    Exit Function

Erro_Documento_Lock:

    Documento_Lock = gErr

    Select Case gErr

        Case 98725
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98726, 98727
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO1", gErr, objDoc.iCodigo)

        Case 98728
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_CADASTRADO", gErr, objDoc.iCodigo)

        Case 98729
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_DOCUMENTO", gErr, objDoc.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItemServico_Lock2(ByVal objCompServ As ClassCompServ) As Long
'Rotina executada dentro de uma TRANSACAO!!!!!!!!!!!!
'Faz lock shared dos Itens do CompServ

Dim lErro As Long, lComando As Long
Dim iCodDummy As Integer
Dim objItem As ClassCompServItem
Dim objDoc As New ClassDocumento

On Error GoTo Erro_ItemServico_Lock2
    
    lComando = Comando_Abrir
    If lComando = 0 Then gError 98723
    
    For Each objItem In objCompServ.colCompServItem
        
        'Pesquisa o Item no BD
        lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM ItemServico WHERE Codigo = ?", iCodDummy, objItem.iCodItemServico)
        If lErro <> AD_SQL_SUCESSO Then gError 98647

        'Se posiciona no registro que vai receber o lock
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98648

        'Se não encontrou o Item --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 98649
        
        'Locka efetivamente o Item
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 98650
       
        'esses 2 ifs leem o codigo do doc a partir do nome reduzido e fazem
        'o lock do mesmo
        If Len(Trim(objItem.sDocExtTipo)) > 0 Then
            objDoc.sNomeReduzido = objItem.sDocExtTipo
            
            lErro = Documento_Lock(objDoc)
            If lErro <> SUCESSO Then gError 98730
            objItem.iDocExtTipo = objDoc.iCodigo
        End If
       
        If Len(Trim(objItem.sDocIntTipo)) > 0 Then
            objDoc.sNomeReduzido = objItem.sDocIntTipo
            
            lErro = Documento_Lock(objDoc)
            If lErro <> SUCESSO Then gError 98731
            objItem.iDocIntTipo = objDoc.iCodigo
        End If
    
    Next
    
    Call Comando_Fechar(lComando)
    
    ItemServico_Lock2 = SUCESSO

    Exit Function

Erro_ItemServico_Lock2:

    ItemServico_Lock2 = gErr

    Select Case gErr

        Case 98647, 98648
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO1", gErr)

        Case 98649
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_ITEMSERVICO", gErr, objItem.iCodItemServico, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 98650
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMSERVICO", gErr, objItem.iCodItemServico)
        
        Case 98723
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98730, 98731
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Verifica_Cliente1(ByVal lCodigoComp, ByVal lCodigoCli As Long) As Long
'Faz a verificação da validade do comprovante com o cliente

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_Verifica_Cliente1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99168
    
    'Lê os dados da tabela TabPrecoItensGR com o produto passado, verificando se combina
    'com a tabela de preço e data de vigencia em questão
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM CompServGR, SolicitacaoServico WHERE Codigo =? AND SolicitacaoServico.NumIntDoc = CompServGR.NumIntDocOrigem AND Cliente = ? ", lCodigo, lCodigoComp, lCodigoCli)
    If lErro <> AD_SQL_SUCESSO Then gError 99169
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99170
    
    'Se não encontrar --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 99171
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Verifica_Cliente1 = SUCESSO

    Exit Function

Erro_Verifica_Cliente1:

    Verifica_Cliente1 = gErr

    Select Case gErr
        
        Case 99168
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 99169, 99170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPROVANTESSERVICOS", gErr, lCodigoComp)
        
        Case 99171
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function
   
End Function

Public Function CompServFrete_Le(ByVal objCompServ As ClassCompServ) As Long
'Lê o Comprovante Serv numa view para a nota fiscal
'Código e FilialEmpresa passados como parametros

Dim lErro As Long, lComando As Long
Dim tCompServ As typeCompServ
Dim lCliente As Long

On Error GoTo Erro_CompServFrete_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99163

    'inicializando
    tCompServ.sCodigoContainer = String(STRING_CODCONTAINER, 0)
    tCompServ.sLacre = String(STRING_LACRE, 0)
    tCompServ.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServ.sProduto = String(STRING_PRODUTO_CODIGO, 0)
    tCompServ.sUM = String(STRING_UM_NOME, 0)

    'Busca na view ComprovServFrete a Cotacao com Código e FilialEmpresa informados
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, FilialEmpresa, Codigo, DataEmissao, Produto, QuantMaterial, UM, ValorMercadoria," _
    & "FretePeso, Pedagio, AdValoren, DataDemurrage, CodigoContainer, ValorContainer, Tara, Lacre, Observacao, NumIntNota, Situacao, Quantidade, ClienteSol " _
    & "FROM CompServFrete " _
    & "WHERE Codigo=? AND FilialEmpresa=?", tCompServ.lNumIntDoc, tCompServ.lNumIntDocOrigem, tCompServ.iFilialEmpresa, tCompServ.lCodigo, tCompServ.dtDataEmissao, _
    tCompServ.sProduto, tCompServ.dQuantMaterial, tCompServ.sUM, tCompServ.dValorMercadoria, tCompServ.dFretePeso, tCompServ.dPedagio, _
    tCompServ.dAdValoren, tCompServ.dtDataDemurrage, tCompServ.sCodigoContainer, tCompServ.dValorContainer, tCompServ.dTara, tCompServ.sLacre, tCompServ.sObservacao, tCompServ.lNumIntNota, tCompServ.iSituacao, tCompServ.dQuantidade, lCliente, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99164

    'Se posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99165

    'Se não encontrou o CompServ
    If lErro = AD_SQL_SEM_DADOS Then gError 99166

    'Carrega o obj com o conteudo lido do bd
    With objCompServ

        .dAdValoren = tCompServ.dAdValoren
        .dFretePeso = tCompServ.dFretePeso
        .dPedagio = tCompServ.dPedagio
        .dQuantMaterial = tCompServ.dQuantMaterial
        .dTara = tCompServ.dTara
        .dtDataDemurrage = tCompServ.dtDataDemurrage
        .dtDataEmissao = tCompServ.dtDataEmissao
        .dValorMercadoria = tCompServ.dValorMercadoria
        .lNumIntDoc = tCompServ.lNumIntDoc
        .lNumIntDocOrigem = tCompServ.lNumIntDocOrigem
        .sCodigoContainer = tCompServ.sCodigoContainer
        .dValorContainer = tCompServ.dValorContainer
        .sLacre = tCompServ.sLacre
        .sObservacao = tCompServ.sObservacao
        .sProduto = tCompServ.sProduto
        .sUM = tCompServ.sUM
        .iSituacao = tCompServ.iSituacao
        .lNumIntNota = tCompServ.lNumIntNota
        .dQuantidade = tCompServ.dQuantidade
        .lCliente = lCliente
        
    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServFrete_Le = SUCESSO

    Exit Function

Erro_CompServFrete_Le:

    CompServFrete_Le = gErr

    Select Case gErr

        Case 99163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99164, 99165
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV2", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 99166
            'NAO ACHOU COMPSERV !!

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CompServGR_Le_Frete(ByVal objCompServ As ClassCompServ) As Long
'Lê o Comprovante Serv. a partir do
'Código e FilialEmpresa passados como parametros

Dim lErro As Long, lComando As Long
Dim tCompServ As typeCompServ
Dim lCliente As Long
Dim iTipoOperacao As Integer

On Error GoTo Erro_CompServGR_Le_Frete

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99313

    'inicializando
    tCompServ.sCodigoContainer = String(STRING_CODCONTAINER, 0)
    tCompServ.sLacre = String(STRING_LACRE, 0)
    tCompServ.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCompServ.sProduto = String(STRING_PRODUTO_CODIGO, 0)
    tCompServ.sUM = String(STRING_UM_NOME, 0)
    tCompServ.sEmbalagem = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tCompServ.sCidadeDestino = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    tCompServ.sCidadeOrigem = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)
    tCompServ.sUFDestino = String(STRING_ORIGEMDESTINO_UF, 0)
    tCompServ.sUFOrigem = String(STRING_ORIGEMDESTINO_UF, 0)
    tCompServ.sMaterial = String(STRING_SOLICITACAOSERVICO_MATERIAL, 0)
    'Busca na tabela CompServGR a Cotacao com Código e FilialEmpresa informados
    'e armazena os dados no type
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEmissao, Produto, QuantMaterial, UM, ValorMercadoria," _
    & "FretePeso, Pedagio, AdValoren, DataDemurrage, CodigoContainer, ValorContainer, Tara, Lacre, NumIntNota, Situacao, Quantidade, DescEmbalagem, CidadeOrigem, UFOrigem, CidadeDestino, UFDestino,Material " _
    & "FROM CompSolFrete " _
    & "WHERE Codigo=? AND FilialEmpresa=?", tCompServ.lNumIntDoc, tCompServ.dtDataEmissao, tCompServ.sProduto, tCompServ.dQuantMaterial, tCompServ.sUM, tCompServ.dValorMercadoria, tCompServ.dFretePeso, tCompServ.dPedagio, _
    tCompServ.dAdValoren, tCompServ.dtDataDemurrage, tCompServ.sCodigoContainer, tCompServ.dValorContainer, tCompServ.dTara, tCompServ.sLacre, tCompServ.lNumIntNota, tCompServ.iSituacao, tCompServ.dQuantidade, tCompServ.sEmbalagem, tCompServ.sCidadeOrigem, tCompServ.sUFOrigem, tCompServ.sCidadeDestino, tCompServ.sUFDestino, tCompServ.sMaterial, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99314

    'Se posiciona no registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99315

    'Se não encontrou o CompServ ==> erro 99316
    If lErro = AD_SQL_SEM_DADOS Then gError 99316

    'Carrega o obj com o conteudo lido do bd
    With objCompServ

        .dAdValoren = tCompServ.dAdValoren
        .dFretePeso = tCompServ.dFretePeso
        .dPedagio = tCompServ.dPedagio
        .dQuantMaterial = tCompServ.dQuantMaterial
        .dTara = tCompServ.dTara
        .dtDataDemurrage = tCompServ.dtDataDemurrage
        .dtDataEmissao = tCompServ.dtDataEmissao
        .dValorMercadoria = tCompServ.dValorMercadoria
        .lNumIntDoc = tCompServ.lNumIntDoc
        .sCodigoContainer = tCompServ.sCodigoContainer
        .dValorContainer = tCompServ.dValorContainer
        .sLacre = tCompServ.sLacre
        .sObservacao = tCompServ.sObservacao
        .sProduto = tCompServ.sProduto
        .sUM = tCompServ.sUM
        .iSituacao = tCompServ.iSituacao
        .lNumIntNota = tCompServ.lNumIntNota
        .dQuantidade = tCompServ.dQuantidade
        .sEmbalagem = tCompServ.sEmbalagem
        .sCidadeOrigem = tCompServ.sCidadeOrigem
        .sUFDestino = tCompServ.sUFDestino
        .sUFOrigem = tCompServ.sUFOrigem
        .sCidadeDestino = tCompServ.sCidadeDestino
        .sMaterial = tCompServ.sMaterial
    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    CompServGR_Le_Frete = SUCESSO

    Exit Function

Erro_CompServGR_Le_Frete:

    CompServGR_Le_Frete = gErr

    Select Case gErr

        Case 99313
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99314, 99315
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV2", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 99316
            'NAO ACHOU COMPSERV !!

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function SolicitacaoServicoNaoAtendidos_Le(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Le os dados da tabela Solicitação de Serviço

Dim lErro As Long
Dim TypeSolServ As typeSolicitacaoServico
Dim sDespachante As String
Dim sCliente As String
Dim lComando As Long

On Error GoTo Erro_SolicitacaoServicoNaoAtendidos_Le
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99319
    
    With TypeSolServ
        
        'Inicializar as variáveis
        .sNumReferencia = String(STRING_SOLICITACAOSERVICO_NUMREFERENCIA, 0)
        .sMaterial = String(STRING_SOLICITACAOSERVICO_MATERIAL, 0)
        .sUM = String(STRING_SOLICITACAOSERVICO_UM, 0)
        .sBooking = String(STRING_SOLICITACAOSERVICO_BOOKING, 0)
        .sObservacao = String(STRING_SOLICITACAOSERVICO_OBSERVACAO, 0)
        sDespachante = String(STRING_DESPACHANTE_NOMEREDUZIDO, 0)
        sCliente = String(STRING_CLIENTES_NOMEREDUZIDO, 0)
        
        'Le a View SolicitacaoRel para recolher do Banco todos os dados da tabela SolicitacaoServico, o nome reduzido do Cliente e do Despachante.
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, SolicitacaoRel.CodTabPreco, Cliente, DataEmissao, NumReferencia, DataPedido, TipoOperacao, Despachante, Material, QuantMaterial, UM, SolicitacaoRel.ValorMercadoria, TipoEmbalagem, TipoContainer, CodProgNavio, Booking, DataPrevInicio, HoraPrevInicio, DataPrevFim, HoraPrevFim, Observacao, EnderecoOrigem, EnderecoDestino, " _
        & "CliNomeRed , DesNomeRed FROM SolicitacaoRel, SolServServicoNaoAtendidos Where SolservServicoNaoAtendidos.NumIntDocOrigem = SolicitacaoRel.NumIntDoc AND SolicitacaoRel.Numero = ? AND SolicitacaoRel.FilialEmpresa = ?", .lNumIntDoc, .lCodTabPreco, .lCliente, .dtDataEmissao, .sNumReferencia, .dtDataPedido, .iTipoOperacao, .iDespachante, .sMaterial, .dQuantMaterial, .sUM, .dValorMercadoria, .iTipoEmbalagem, .iTipoContainer, .lCodProgNavio, .sBooking, .dtDataPrevInicio, .dHoraPrevInicio, _
        .dtDataPrevFim, .dHoraPrevFim, .sObservacao, .lEnderecoOrigem, .lEnderecoDestino, sCliente, sDespachante, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 99320
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99321
        
        'Se não encontrou --> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 99322
        
        'Inicializa os objEndereco
        Set objSolicitacaoServico.objEnderecoOrigem = New ClassEndereco
        Set objSolicitacaoServico.objEnderecoDestino = New ClassEndereco
        
        'Carrega o obj com os endereços recolhidos no banco
        objSolicitacaoServico.objEnderecoOrigem.lCodigo = .lEnderecoOrigem
        objSolicitacaoServico.objEnderecoDestino.lCodigo = .lEnderecoDestino
        
        'Lê o endereço com o código passado
        lErro = CF("Endereco_Le", objSolicitacaoServico.objEnderecoOrigem)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 99323
        
        'Se não encontrou o endereço --> erro
        If lErro = 12309 Then gError 99324
        
        'Lê o endereço com o código passado
        lErro = CF("Endereco_Le", objSolicitacaoServico.objEnderecoDestino)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 99325
        
        'Se não encontrou o endereço --> erro
        If lErro = 12309 Then gError 99326
        
        'Carrega o objSolicitacaoServico
        objSolicitacaoServico.dQuantMaterial = .dQuantMaterial
        objSolicitacaoServico.dtDataEmissao = .dtDataEmissao
        objSolicitacaoServico.dtDataPedido = .dtDataPedido
        objSolicitacaoServico.dtDataPrevFim = .dtDataPrevFim
        objSolicitacaoServico.dtDataPrevInicio = .dtDataPrevInicio
        objSolicitacaoServico.dtHoraPrevFim = .dHoraPrevFim
        objSolicitacaoServico.dtHoraPrevInicio = .dHoraPrevInicio
        objSolicitacaoServico.dValorMercadoria = .dValorMercadoria
        objSolicitacaoServico.iTipoContainer = .iTipoContainer
        objSolicitacaoServico.iTipoEmbalagem = .iTipoEmbalagem
        objSolicitacaoServico.iTipoOperacao = .iTipoOperacao
        objSolicitacaoServico.lCodProgNavio = .lCodProgNavio
        objSolicitacaoServico.lCodTabPreco = .lCodTabPreco
        objSolicitacaoServico.lNumIntDoc = .lNumIntDoc
        objSolicitacaoServico.sBooking = .sBooking
        objSolicitacaoServico.sMaterial = .sMaterial
        objSolicitacaoServico.sNumReferencia = .sNumReferencia
        objSolicitacaoServico.sObservacao = .sObservacao
        objSolicitacaoServico.sUM = .sUM
        
    End With
    
    'Carrega o obj com os outros campos que não fazem parte da tabela solicitação
    objSolicitacaoServico.sClienteNomeRed = sCliente
    objSolicitacaoServico.sDespachanteNomeRed = sDespachante
   
    Call Comando_Fechar(lComando)
    
    SolicitacaoServicoNaoAtendidos_Le = SUCESSO

    Exit Function

Erro_SolicitacaoServicoNaoAtendidos_Le:

    SolicitacaoServicoNaoAtendidos_Le = gErr

    Select Case gErr

        Case 99320, 99321
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
            
        Case 99322
        
        Case 99319
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 99323, 99324, 99325, 99326
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Verifica_Serv_SolServ(objSolServ As ClassSolicitacaoServico, objProd As ClassProduto) As Long
'verifica se o servico esta relacionado com a solicitacao
Dim lErro As Long, lComando As Long
Dim sProd As String

On Error GoTo Erro_Verifica_Serv_SolServ

    'inicializa a string
    sProd = String(STRING_PRODUTO, 0)
    
    'abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98734

    lErro = Comando_Executar(lComando, "SELECT Codigo from Produtos WHERE Codigo = ? AND Codigo in (SELECT produto FROM SolServServicoNaoAtendidos WHERE numintdocorigem = ?)", sProd, objProd.sCodigo, objSolServ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98736

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98737
    
    If lErro <> AD_SQL_SUCESSO Then gError 98738

    'fecha comando
    Call Comando_Fechar(lComando)
    
    Verifica_Serv_SolServ = SUCESSO
    
    Exit Function
    
Erro_Verifica_Serv_SolServ:

    Verifica_Serv_SolServ = gErr

    Select Case gErr
    
        Case 98734
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98736, 98737
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO1", gErr)
        
        Case 98738
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_RELACIONADO_COM_SOLSERV", gErr, objProd.sCodigo, objSolServ.lNumero, objSolServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function
