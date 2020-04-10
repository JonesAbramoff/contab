VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOMSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function Requisitante_Le(ByVal objRequisitante As ClassRequisitante) As Long
'Lê os dados do requisitante, cujo código foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tRequisitante As typeRequisitante
Dim sCcl As String
Dim sNome As String
Dim sNomeReduzido As String

On Error GoTo Erro_Requisitante_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49081

    tRequisitante.sNome = String(STRING_REQUISITANTE_NOME, 0)
    tRequisitante.sNomeReduzido = String(STRING_REQUISITANTE_NOMERED, 0)
    tRequisitante.sCcl = String(STRING_REQUISITANTE_CCL, 0)
    tRequisitante.sEmail = String(64, 0)
    tRequisitante.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Lê os dados da tabela Requisitante
    lErro = Comando_Executar(lComando, "SELECT  Nome, NomeReduzido, Ccl, Email, CodUsuario FROM Requisitante WHERE Codigo = ? ", tRequisitante.sNome, tRequisitante.sNomeReduzido, tRequisitante.sCcl, tRequisitante.sEmail, tRequisitante.sCodUsuario, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49082

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49083

    'Requisitante nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 49084

    'carrega os dados em objRequisitante
    objRequisitante.sNome = tRequisitante.sNome
    objRequisitante.sNomeReduzido = tRequisitante.sNomeReduzido
    objRequisitante.sCcl = tRequisitante.sCcl
    objRequisitante.sEmail = tRequisitante.sEmail
    objRequisitante.sCodUsuario = tRequisitante.sCodUsuario

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Requisitante_Le = SUCESSO

    Exit Function

Erro_Requisitante_Le:

    Requisitante_Le = Err

    Select Case Err

        Case 49081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49082, 49083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE", Err, objRequisitante.lCodigo)

        Case 49084

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147582)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Requisitante_Le_NomeReduzido(ByVal objRequisitante As ClassRequisitante) As Long
'Lê os dados do requisitante, cujo nome reduzido foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tRequisitante As typeRequisitante

On Error GoTo Erro_Requisitante_Le_NomeReduzido

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51149

    tRequisitante.sNome = String(STRING_REQUISITANTE_NOME, 0)
    tRequisitante.sCcl = String(STRING_REQUISITANTE_CCL, 0)
    tRequisitante.sEmail = String(64, 0)
    tRequisitante.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    'Lê os dados da tabela Requisitante baseado no Nome Reduzido
    lErro = Comando_Executar(lComando, "SELECT  Codigo,Nome,Ccl, Email, CodUsuario FROM Requisitante WHERE NomeReduzido = ? ", tRequisitante.lCodigo, tRequisitante.sNome, tRequisitante.sCcl, tRequisitante.sEmail, tRequisitante.sCodUsuario, objRequisitante.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 51150

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51151
    If lErro = AD_SQL_SEM_DADOS Then Error 51152 'Não encontrou

    'carrega os dados em objRequisitante
    objRequisitante.lCodigo = tRequisitante.lCodigo
    objRequisitante.sNome = tRequisitante.sNome
    objRequisitante.sCcl = tRequisitante.sCcl
    objRequisitante.sEmail = tRequisitante.sEmail
    objRequisitante.sCodUsuario = tRequisitante.sCodUsuario

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Requisitante_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Requisitante_Le_NomeReduzido:

    Requisitante_Le_NomeReduzido = Err

    Select Case Err

        Case 51149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 51150, 51151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE1", Err, objRequisitante.sNomeReduzido)

        Case 51152 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147583)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BaixaReqCompras_ObterRequisicoes(ByVal objBaixaReqCompras As ClassBaixaReqCompras) As Long
'Obtem as Requisicoes de Compras a Baixar de acordo
'com a seleção determinada na tela

Dim lErro As Long
Dim sSelect As String
Dim lNumIntDocRC As Long
Dim sCodigoOP As String
Dim iFilialEmpresaRC As Integer
Dim lCodigoRC As Long
Dim dtDataRC As Date
Dim dtDataLimite As Date
Dim lRequisitanteRC As Long
Dim sCcl As String
Dim dPercentualRecebido As Double
Dim sNomeReduzido As String
Dim dResiduo As Double
Dim alComando(0 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_BaixaReqCompras_ObterRequisicoes

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51166
    Next

    'Montar SELECT das Requisições dinamicamente
    Call BaixaReqCompras_ObterRequisicoes1(objBaixaReqCompras, sSelect)

    sCcl = String(STRING_CCL, 0)
    sNomeReduzido = String(STRING_REQUISITANTE_NOMERED, 0)
    dtDataRC = DATA_NULA
    dtDataLimite = DATA_NULA
    sCodigoOP = String(STRING_OPCODIGO, 0)

    'preparar parte fixa do SELECT
    lErro = BaixaReqCompras_ObterRequisicoes2(alComando(0), sSelect, sCodigoOP, iFilialEmpresaRC, dtDataRC, dtDataLimite, lCodigoRC, lRequisitanteRC, sCcl, sNomeReduzido, dPercentualRecebido, lNumIntDocRC)
    If lErro <> SUCESSO Then Error 51167

    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das requisições
    lErro = BaixaReqCompras_ObterRequisicoes3(alComando(0), objBaixaReqCompras)
    If lErro <> SUCESSO Then Error 63358

    'Processar os registros lidos
    lErro = BaixaReqCompras_ObterRequisicoes4(alComando(0), objBaixaReqCompras, sCodigoOP, iFilialEmpresaRC, dtDataRC, dtDataLimite, lCodigoRC, lRequisitanteRC, sCcl, sNomeReduzido, dPercentualRecebido, lNumIntDocRC)
    If lErro <> SUCESSO And lErro <> 63371 Then Error 63359

    If lErro = 63371 Then Error 63406
    
    'Verifica se quer trazer para a tela apenas as Requisicoes residuais
    If objBaixaReqCompras.iSoResiduais = 1 Then
        '???? A descrição da função não está correspondento ao que ela significa.
        'exatamente ao que ela se propõe. (ex.: BaixaReqCompras_Obter_Residuais)
        'Lê os Itens de Requisições de Compras que são Residuais e adiciona-os em objBaixaReqCompras.colReqComprasInfo
        lErro = BaixaReqCompras_Obter_Residuais(objBaixaReqCompras)
        If lErro <> SUCESSO Then Error 63424

    End If


    'Fechar Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    BaixaReqCompras_ObterRequisicoes = SUCESSO

    Exit Function

Erro_BaixaReqCompras_ObterRequisicoes:

    BaixaReqCompras_ObterRequisicoes = Err

    Select Case Err

        Case 51166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 51167, 63358, 63359
            'Erros tratados nas rotinas chamadas

        Case 63406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENCONTRADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147584)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Sub BaixaReqCompras_ObterRequisicoes1(ByVal objBaixaReqCompras As ClassBaixaReqCompras, sSelect As String)
'Montar SELECT das Requisições dinamicamente

Dim sFields As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sGroupBy As String
Dim sHaving As String

    sSelect = "SELECT MIN(RequisicaoCompra.OPCodigo) AS OPCodigo ,  MIN(RequisicaoCompra.FilialEmpresa)AS FilialEmpresa, " _
              & "MIN(RequisicaoCompra.Data) AS Data, MIN(RequisicaoCompra.DataLimite)AS DataLimite," _
              & "MIN(RequisicaoCompra.Codigo) AS Codigo, MIN(RequisicaoCompra.Requisitante) AS Requisitante," _
              & "MIN(RequisicaoCompra.Ccl) AS Ccl, MIN(Requisitante.NomeReduzido) AS NomeReduzido, " _
              & "MIN(ItensReqCompra.QuantRecebida/(ItensReqCompra.Quantidade-ItensReqCompra.QuantCancelada)) AS PercReceb,RequisicaoCompra.NumIntDoc"
    sFrom = " FROM RequisicaoCompra, ItensReqCompra, Requisitante"
'    sWhere = " WHERE (RequisicaoCompra.DataEnvio <> ?) AND (RequisicaoCompra.NumIntDoc=ItensReqCompra.ReqCompra)AND Requisitante.Codigo=RequisicaoCompra.Requisitante AND ItensReqCompra.Status <> ? "
    sWhere = " WHERE (ItensReqCompra.Quantidade<>ItensReqCompra.QuantCancelada ) AND (RequisicaoCompra.DataEnvio <> ?) AND (RequisicaoCompra.NumIntDoc=ItensReqCompra.ReqCompra)AND Requisitante.Codigo=RequisicaoCompra.Requisitante AND ItensReqCompra.Status <> ? "
    sWhere = sWhere & " AND RequisicaoCompra.NumIntDoc NOT IN (((SELECT ItensReqCompra.ReqCompra FROM ItensReqCompra, ItemRCItemPC, ItensPedCompra "
    sWhere = sWhere & " WHERE ItensReqCompra.NumIntDoc = ItemRCItemPC.ItemRC AND ItemRCItemPC.ItemPC = ItensPedCompra.NumIntDoc) UNION (SELECT ItensReqCompra.ReqCompra FROM ItensReqCompra, ItensConcorrencia, ItemRCItemConcorrencia"
    sWhere = sWhere & " WHERE ItensReqCompra.NumIntDoc = ItemRCItemConcorrencia.ItemReqCompra AND ItemRCItemConcorrencia.ItemConcorrencia = ItensConcorrencia.NumIntDoc)))"
    
    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND RequisicaoCompra.FilialEmpresa = ?"
    If objBaixaReqCompras.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data >=?"
    If objBaixaReqCompras.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data <=?"
    If objBaixaReqCompras.dtDataLimiteDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite >=?"
    If objBaixaReqCompras.dtDataLimiteAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite <=?"
    If objBaixaReqCompras.lRequisicaoDe <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo >= ? "
    If objBaixaReqCompras.lRequisicaoAte <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo <= ? "
    If objBaixaReqCompras.lRequisitanteDe <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Requisitante >=? "
    If objBaixaReqCompras.lRequisitanteAte <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Requisitante <=? "
    If Len(Trim(objBaixaReqCompras.sCclDe)) > 0 Then sWhere = sWhere & " AND RequisicaoCompra.Ccl >= ?"
    If Len(Trim(objBaixaReqCompras.sCclAte)) > 0 Then sWhere = sWhere & " AND RequisicaoCompra.Ccl <= ?"
    
    sGroupBy = " GROUP BY " & objBaixaReqCompras.sOrdenacao & ", RequisicaoCompra.NumIntDoc"
    sOrderBy = " ORDER BY " & objBaixaReqCompras.sOrdenacao & ", RequisicaoCompra.NumIntDoc"
    sHaving = ""
    sSelect = sSelect & sFrom & sWhere & sGroupBy & sOrderBy & sHaving

    Exit Sub

End Sub

Function BaixaReqCompras_ObterRequisicoes2(ByVal lComando As Long, ByVal sSelect As String, vsCodigoOP As Variant, viFilialEmpresaRC As Variant, vdtDataRC As Variant, vdtDataLimite As Variant, vlCodigoRC As Variant, vlRequisitanteRC As Variant, vsCcl As Variant, vsNomeReduzido As Variant, vdPercentualRecebido As Variant, vlNumIntDocRC As Variant) As Long

Dim ret As Integer, lErro As Long

On Error GoTo Erro_BaixaReqCompras_ObterRequisicoes2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 51168

    ret = Comando_BindVarInt(lComando, vsCodigoOP)
    If (ret <> AD_SQL_SUCESSO) Then gError 178854

    ret = Comando_BindVarInt(lComando, viFilialEmpresaRC)
    If (ret <> AD_SQL_SUCESSO) Then gError 51170

    ret = Comando_BindVarInt(lComando, vdtDataRC)
    If (ret <> AD_SQL_SUCESSO) Then gError 51171

    ret = Comando_BindVarInt(lComando, vdtDataLimite)
    If (ret <> AD_SQL_SUCESSO) Then gError 51172

    ret = Comando_BindVarInt(lComando, vlCodigoRC)
    If (ret <> AD_SQL_SUCESSO) Then gError 51173

    ret = Comando_BindVarInt(lComando, vlRequisitanteRC)
    If (ret <> AD_SQL_SUCESSO) Then gError 51174

    ret = Comando_BindVarInt(lComando, vsCcl)
    If (ret <> AD_SQL_SUCESSO) Then gError 51175

    ret = Comando_BindVarInt(lComando, vsNomeReduzido)
    If (ret <> AD_SQL_SUCESSO) Then gError 51176

    ret = Comando_BindVarInt(lComando, vdPercentualRecebido)
    If (ret <> AD_SQL_SUCESSO) Then gError 62535

    ret = Comando_BindVarInt(lComando, vlNumIntDocRC)
    If (ret <> AD_SQL_SUCESSO) Then gError 51169

    BaixaReqCompras_ObterRequisicoes2 = SUCESSO

    Exit Function

Erro_BaixaReqCompras_ObterRequisicoes2:

    BaixaReqCompras_ObterRequisicoes2 = gErr

    Select Case gErr

        Case 51168 To 51176, 62535, 178854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147585)

    End Select

    Exit Function

End Function


Function BaixaReqCompras_ObterRequisicoes3(ByVal lComando As Long, ByVal objBaixaReqCompras As ClassBaixaReqCompras) As Long

Dim lErro As Long
Dim lNumIntDocRC As Variant
Dim viFilialEmpresaRC As Variant
Dim lCodigoRC As Variant
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant
Dim vdtDataLimiteDe As Variant
Dim vdtDataLimiteAte As Variant
Dim vlRequisitanteDe As Variant
Dim vlRequisitanteAte As Variant
Dim vsCclDe As Variant
Dim vsCclAte As Variant
Dim vlRequisicaoDe As Variant
Dim vlRequisicaoAte As Variant
Dim vdtDataNula As Variant
Dim dPercentualRecebido As Variant
Dim vStatus As Variant
Dim sNomeReduzido As Variant

On Error GoTo Erro_BaixaReqCompras_ObterRequisicoes3

    vdtDataNula = DATA_NULA
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then Error 63360

    vStatus = CInt(STATUS_BAIXADO)
    lErro = Comando_BindVarInt(lComando, vStatus)
    If (lErro <> AD_SQL_SUCESSO) Then Error 63364

    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresaRC = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresaRC)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63360
    End If

    If objBaixaReqCompras.dtDataDe <> DATA_NULA Then
        vdtDataDe = objBaixaReqCompras.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63364
    End If

    If objBaixaReqCompras.dtDataAte <> DATA_NULA Then
        vdtDataAte = objBaixaReqCompras.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63363
    End If

    If objBaixaReqCompras.dtDataLimiteDe <> DATA_NULA Then
        vdtDataLimiteDe = objBaixaReqCompras.dtDataLimiteDe
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63366
    End If

    If objBaixaReqCompras.dtDataLimiteAte <> DATA_NULA Then
        vdtDataLimiteAte = objBaixaReqCompras.dtDataLimiteAte
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63365
    End If

    If objBaixaReqCompras.lRequisicaoDe <> 0 Then
        vlRequisicaoDe = objBaixaReqCompras.lRequisicaoDe
        lErro = Comando_BindVarInt(lComando, vlRequisicaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63361
    End If

    If objBaixaReqCompras.lRequisicaoAte <> 0 Then
        vlRequisicaoAte = objBaixaReqCompras.lRequisicaoAte
        lErro = Comando_BindVarInt(lComando, vlRequisicaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63362
    End If

    If objBaixaReqCompras.lRequisitanteDe <> 0 Then
        vlRequisitanteDe = objBaixaReqCompras.lRequisitanteDe
        lErro = Comando_BindVarInt(lComando, vlRequisitanteDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63368
    End If

    If objBaixaReqCompras.lRequisitanteAte <> 0 Then
        vlRequisitanteAte = objBaixaReqCompras.lRequisitanteAte
        lErro = Comando_BindVarInt(lComando, vlRequisitanteAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63367
    End If

    If Len(Trim(objBaixaReqCompras.sCclDe)) <> 0 Then
        vsCclDe = objBaixaReqCompras.sCclDe
        lErro = Comando_BindVarInt(lComando, vsCclDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63374
    End If

    If Len(Trim(objBaixaReqCompras.sCclAte)) <> 0 Then
        vsCclAte = objBaixaReqCompras.sCclAte
        lErro = Comando_BindVarInt(lComando, vsCclAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 63375
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 63369

    BaixaReqCompras_ObterRequisicoes3 = SUCESSO

    Exit Function

Erro_BaixaReqCompras_ObterRequisicoes3:

    BaixaReqCompras_ObterRequisicoes3 = Err

    Select Case Err

        Case 63360 To 63369, 63374, 63375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICOES_COMPRA_BAIXA_RC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147586)

    End Select

    Exit Function

End Function

Function BaixaReqCompras_ObterRequisicoes4(ByVal lComando As Long, ByVal objBaixaReqCompras As ClassBaixaReqCompras, sCodigoOP As String, iFilialEmpresaRC As Integer, dtDataRC As Date, dtDataLimite As Date, lCodigoRC As Long, lRequisitanteRC As Long, sCcl As String, sNomeReduzido As String, dPercentualRecebido As Double, lNumIntDocRC As Long) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objReqComprasInfo As ClassReqComprasInfo
Dim iContador As Integer

On Error GoTo Erro_BaixaReqCompras_ObterRequisicoes4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63370

    'Se não encontrou==>Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 63371

    iContador = 0

    Do While lErro = AD_SQL_SUCESSO

        iContador = iContador + 1

        'Verifica se o número de requisicoes é superior ao número máximo de requisicoes permitido
        If iContador = NUM_MAX_REQUISICOES Then Error 63372

        Set objReqComprasInfo = New ClassReqComprasInfo

        'Preenche objRequisicaoCompras com os dados lidos do BD
        objReqComprasInfo.dtData = dtDataRC
        objReqComprasInfo.dtDataLimite = dtDataLimite
        objReqComprasInfo.iFilialEmpresa = iFilialEmpresaRC
        objReqComprasInfo.lCodRequisicao = lCodigoRC
        objReqComprasInfo.lNumIntReq = lNumIntDocRC
        objReqComprasInfo.lRequisitante = lRequisitanteRC
        objReqComprasInfo.dMinPercRecItens = dPercentualRecebido
        objReqComprasInfo.sCcl = sCcl
        objReqComprasInfo.sNomeRequisitante = sNomeReduzido
        objReqComprasInfo.sCodigoOP = sCodigoOP
        
        'Adiciona objReqComprasInfo na coleção de Requisicoes de Compra
        objBaixaReqCompras.colReqComprasInfo.Add objReqComprasInfo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63373

    Loop

    BaixaReqCompras_ObterRequisicoes4 = SUCESSO

    Exit Function

Erro_BaixaReqCompras_ObterRequisicoes4:

    BaixaReqCompras_ObterRequisicoes4 = Err

    Select Case Err

        Case 63370, 63373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", Err, objReqComprasInfo.lCodRequisicao)

        Case 63371
            'Erro tratado na rotina chamadora

        Case 63372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147587)

    End Select

    Exit Function

End Function


Function LiberacaoDeBloqueiosPC_ObterBloqueios(ByVal objLiberacaoBloqueiosPC As ClassLiberacaoBloqueiosPC, ByVal colTiposBloqueio As Collection) As Long
'Preenche colBloqueioLiberacaoInfo em objLiberacaoBloqueiosPC

Dim lErro As Long, sSelect As String
Dim lComando As Long
Dim iBloqueiosPCFilialEmpresa As Integer, lBloqueiosPCPedCompras As Long, iBloqueiosPCSequencial As Integer, sBloqueiosPCCodUsuario As String, dtBloqueiosPCData As Date, iBloqueiosPCTipoDeBloqueio As Integer, sBloqueiosPCResponsavel As String
Dim dtPedidoCompraDataEmissao As Date, dPedidoCompraValorTotal As Double, sFornecedoresNomeReduzido As String, sTiposDeBloqueioPCNomeReduzido As String, dtDataNula As Date

On Error GoTo Erro_LiberacaoDeBloqueiosPC_ObterBloqueios

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51176

    'Montar SELECT dinamicamente
    Call LiberacaoDeBloqueiosPC_ObterBloqueios1(objLiberacaoBloqueiosPC, sSelect)

    sBloqueiosPCCodUsuario = String(STRING_SIGLA_USUARIO, 0)
    sBloqueiosPCResponsavel = String(STRING_BLOQUEIOSPC_RESPONSAVEL, 0)
    sFornecedoresNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    sTiposDeBloqueioPCNomeReduzido = String(STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, 0)
    dtDataNula = DATA_NULA
    
    'Preparar parte fixa do SELECT
    lErro = LiberacaoDeBloqueiosPC_ObterBloqueios2(lComando, sSelect, iBloqueiosPCFilialEmpresa, lBloqueiosPCPedCompras, iBloqueiosPCSequencial, sBloqueiosPCCodUsuario, dtBloqueiosPCData, iBloqueiosPCTipoDeBloqueio, sBloqueiosPCResponsavel, dtPedidoCompraDataEmissao, dPedidoCompraValorTotal, sFornecedoresNomeReduzido, sTiposDeBloqueioPCNomeReduzido, dtDataNula)
    If lErro <> SUCESSO Then Error 51177

    'Complementa a passagem dos parâmetros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao dos bloqueios
    lErro = LiberacaoDeBloqueiosPC_ObterBloqueios3(lComando, objLiberacaoBloqueiosPC)
    If lErro <> SUCESSO Then Error 51178

    'Processar os registros lidos
    lErro = LiberacaoDeBloqueiosPC_ObterBloqueios4(lComando, objLiberacaoBloqueiosPC, iBloqueiosPCFilialEmpresa, lBloqueiosPCPedCompras, iBloqueiosPCSequencial, sBloqueiosPCCodUsuario, dtBloqueiosPCData, iBloqueiosPCTipoDeBloqueio, sBloqueiosPCResponsavel, dtPedidoCompraDataEmissao, dPedidoCompraValorTotal, sFornecedoresNomeReduzido, sTiposDeBloqueioPCNomeReduzido, colTiposBloqueio)
    If lErro <> SUCESSO And lErro <> 51201 Then Error 51179

    If lErro = 51201 Then Error 51180
    
    Call Comando_Fechar(lComando)

    LiberacaoDeBloqueiosPC_ObterBloqueios = SUCESSO

    Exit Function

Erro_LiberacaoDeBloqueiosPC_ObterBloqueios:

    LiberacaoDeBloqueiosPC_ObterBloqueios = Err

    Select Case Err

        Case 51176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 51177 To 51180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147588)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub LiberacaoDeBloqueiosPC_ObterBloqueios1(ByVal objLiberacaoBloqueiosPC As ClassLiberacaoBloqueiosPC, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String

    sFields = "BloqueiosPC.FilialEmpresa, BloqueiosPC.PedCompras, BloqueiosPC.Sequencial, BloqueiosPC.CodUsuario, BloqueiosPC.Data, BloqueiosPC.TipoDeBloqueio, BloqueiosPC.Responsavel, " & _
                "PedidoCompra.DataEmissao, PedidoCompra.ValorTotal, Fornecedores.NomeReduzido, TiposDeBloqueioPC.NomeReduzido"

    sFrom = " FROM BloqueiosPC, PedidoCompra, Fornecedores, TiposDeBloqueioPC"
    sWhere = " WHERE BloqueiosPC.DataLib = ? AND BloqueiosPC.FilialEmpresa = PedidoCompra.FilialEmpresa AND BloqueiosPC.PedCompras = PedidoCompra.Codigo AND PedidoCompra.Fornecedor = Fornecedores.Codigo AND BloqueiosPC.TipoDeBloqueio = TiposDeBloqueioPC.Codigo"
    sOrderBy = " ORDER BY " & objLiberacaoBloqueiosPC.sOrdenacao

    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND BloqueiosPC.FilialEmpresa = ?"

    If objLiberacaoBloqueiosPC.lPedComprasDe <> 0 Then sWhere = sWhere & " AND PedidoCompra.Codigo >= ?"

    If objLiberacaoBloqueiosPC.lPedComprasAte <> 0 Then sWhere = sWhere & " AND PedidoCompra.Codigo <= ?"

    If objLiberacaoBloqueiosPC.dtBloqueioDe <> DATA_NULA Then sWhere = sWhere & " AND BloqueiosPC.Data >= ?"

    If objLiberacaoBloqueiosPC.dtBloqueioAte <> DATA_NULA Then sWhere = sWhere & " AND BloqueiosPC.Data <= ?"

    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub

Private Function LiberacaoDeBloqueiosPC_ObterBloqueios2(ByVal lComando As Long, ByVal sSelect As String, vBloqueiosPCFilialEmpresa As Variant, vBloqueiosPCPedidoCompras As Variant, vBloqueiosPCSequencial As Variant, vBloqueiosPCCodUsuario As Variant, vBloqueiosPCData As Variant, vBloqueiosPCTipoDeBloqueio As Variant, vBloqueiosPCResponsavel As Variant, vPedidoCompraDataEmissao As Variant, vPedidoCompraValorTotal As Variant, vFornecedoresNomeReduzido As Variant, vTiposDeBloqueioPCNomeReduzido As Variant, vDataNula As Variant) As Long
'Preparar parte fixa do SELECT

Dim ret As Integer, lErro As Long

On Error GoTo Erro_LiberacaoDeBloqueiosPC_ObterBloqueios2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then Error 51181

    ret = Comando_BindVarInt(lComando, vBloqueiosPCFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then Error 51182

    ret = Comando_BindVarInt(lComando, vBloqueiosPCPedidoCompras)
    If (ret <> AD_SQL_SUCESSO) Then Error 51183

    ret = Comando_BindVarInt(lComando, vBloqueiosPCSequencial)
    If (ret <> AD_SQL_SUCESSO) Then Error 51184

    ret = Comando_BindVarInt(lComando, vBloqueiosPCCodUsuario)
    If (ret <> AD_SQL_SUCESSO) Then Error 51185

    ret = Comando_BindVarInt(lComando, vBloqueiosPCData)
    If (ret <> AD_SQL_SUCESSO) Then Error 51186

    ret = Comando_BindVarInt(lComando, vBloqueiosPCTipoDeBloqueio)
    If (ret <> AD_SQL_SUCESSO) Then Error 51187

    ret = Comando_BindVarInt(lComando, vBloqueiosPCResponsavel)
    If (ret <> AD_SQL_SUCESSO) Then Error 51188
    
    ret = Comando_BindVarInt(lComando, vPedidoCompraDataEmissao)
    If (ret <> AD_SQL_SUCESSO) Then Error 51189

    ret = Comando_BindVarInt(lComando, vPedidoCompraValorTotal)
    If (ret <> AD_SQL_SUCESSO) Then Error 51190

    ret = Comando_BindVarInt(lComando, vFornecedoresNomeReduzido)
    If (ret <> AD_SQL_SUCESSO) Then Error 51191

    ret = Comando_BindVarInt(lComando, vTiposDeBloqueioPCNomeReduzido)
    If (ret <> AD_SQL_SUCESSO) Then Error 51192

    ret = Comando_BindVarInt(lComando, vDataNula)
    If (ret <> AD_SQL_SUCESSO) Then Error 51193

    LiberacaoDeBloqueiosPC_ObterBloqueios2 = SUCESSO

    Exit Function

Erro_LiberacaoDeBloqueiosPC_ObterBloqueios2:

    LiberacaoDeBloqueiosPC_ObterBloqueios2 = Err

    Select Case Err

        Case 51181 To 51193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOS_PC_LIBERACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147589)

    End Select

    Exit Function

End Function

Private Function LiberacaoDeBloqueiosPC_ObterBloqueios3(ByVal lComando As Long, ByVal objLiberacaoBloqueiosPC As ClassLiberacaoBloqueiosPC) As Long
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao dos bloqueios

Dim lErro As Long
Dim viFilialEmpresa As Variant, vlPedComprasDe As Variant, vlPedComprasAte As Variant, vdtBloqueioDe As Variant, vdtBloqueioAte As Variant

On Error GoTo Erro_LiberacaoDeBloqueiosPC_ObterBloqueios3

    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then Error 51194
    End If

    If objLiberacaoBloqueiosPC.lPedComprasDe <> 0 Then
        vlPedComprasDe = objLiberacaoBloqueiosPC.lPedComprasDe
        lErro = Comando_BindVarInt(lComando, vlPedComprasDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 51195
    End If

    If objLiberacaoBloqueiosPC.lPedComprasAte <> 0 Then
        vlPedComprasAte = objLiberacaoBloqueiosPC.lPedComprasAte
        lErro = Comando_BindVarInt(lComando, vlPedComprasAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 51196
    End If

    If objLiberacaoBloqueiosPC.dtBloqueioDe <> DATA_NULA Then
        vdtBloqueioDe = objLiberacaoBloqueiosPC.lPedComprasDe
        lErro = Comando_BindVarInt(lComando, vdtBloqueioDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 51197
    End If

    If objLiberacaoBloqueiosPC.dtBloqueioAte <> DATA_NULA Then
        vdtBloqueioAte = objLiberacaoBloqueiosPC.dtBloqueioAte
        lErro = Comando_BindVarInt(lComando, vdtBloqueioAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 51198
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 51199

    LiberacaoDeBloqueiosPC_ObterBloqueios3 = SUCESSO

    Exit Function

Erro_LiberacaoDeBloqueiosPC_ObterBloqueios3:

    LiberacaoDeBloqueiosPC_ObterBloqueios3 = Err

    Select Case Err

        Case 51194 To 51199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOS_PC_LIBERACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147590)

    End Select

    Exit Function

End Function

Private Function LiberacaoDeBloqueiosPC_ObterBloqueios4(ByVal lComando As Long, ByVal objLiberacaoBloqueiosPC As ClassLiberacaoBloqueiosPC, iBloqueiosPCFilialEmpresa As Integer, lBloqueiosPCPedidoCompras As Long, iBloqueiosPCSequencial As Integer, sBloqueiosPCCodUsuario As String, dtBloqueiosPCData As Date, iBloqueiosPCTipoDeBloqueio As Integer, sBloqueiosPCResponsavel As String, dtPedidoCompraDataEmissao As Date, dPedidoCompraValorTotal As Double, sFornecedoresNomeReduzido As String, sTiposDeBloqueioPCNomeReduzido As String, ByVal colTiposBloqueio As Collection) As Long
'Processar os registros lidos

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC, iContador As Integer
Dim iIndice As Integer

On Error GoTo Erro_LiberacaoDeBloqueiosPC_ObterBloqueios4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51200

    If lErro = AD_SQL_SEM_DADOS Then Error 51201

    iContador = 0
    
    Do While lErro = AD_SQL_SUCESSO

        iContador = iContador + 1
        If iContador = NUM_MAX_BLOQUEIOSPC_LIBERACAO Then
        
            '??? chamar rotina aviso p/dizer p/vai trazer apenas n regs
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_NUM_MAX_BLOQUEIOSPC_LIBERACAO")
            
            Exit Do
            
        End If
        
        Set objBloqueioPC = New ClassBloqueioPC
        
        For iIndice = 1 To colTiposBloqueio.Count
        
            If colTiposBloqueio.Item(iIndice) = iBloqueiosPCTipoDeBloqueio Then
        
                objBloqueioPC.iFilialEmpresa = iBloqueiosPCFilialEmpresa
                objBloqueioPC.lPedCompras = lBloqueiosPCPedidoCompras
                objBloqueioPC.iSequencial = iBloqueiosPCSequencial
                objBloqueioPC.sNomeReduzidoFornecedor = sFornecedoresNomeReduzido
                objBloqueioPC.dtDataEmissao = dtPedidoCompraDataEmissao
                objBloqueioPC.dValorPedido = dPedidoCompraValorTotal
                objBloqueioPC.iTipoBloqueio = iBloqueiosPCTipoDeBloqueio
                objBloqueioPC.sNomeReduzidoTipoBloqueio = sTiposDeBloqueioPCNomeReduzido
                objBloqueioPC.sCodUsuario = sBloqueiosPCCodUsuario
                objBloqueioPC.dtData = dtBloqueiosPCData
                objBloqueioPC.sResponsavel = sBloqueiosPCResponsavel
                objBloqueioPC.iMarcado = 1
        
                objLiberacaoBloqueiosPC.colBloqueioPC.Add objBloqueioPC
                
            End If
            
        Next
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51202
            
    Loop

    LiberacaoDeBloqueiosPC_ObterBloqueios4 = SUCESSO

    Exit Function

Erro_LiberacaoDeBloqueiosPC_ObterBloqueios4:

    LiberacaoDeBloqueiosPC_ObterBloqueios4 = Err

    Select Case Err

        Case 51200, 51202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOS_PC_LIBERACAO", Err)

        Case 51201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147591)

    End Select

    Exit Function

End Function

Function GeracaoPedCompraConc_Le_Concorrencias(ByVal objGeracaoPedidoCompra As ClassGeracaoPedCompraConc) As Long
'Pesquisa no BD os dados da concorrência com as características passadas em objGeracaoPedidoCompra

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim lCodigoConc As Long, sDescricaoConc As String, lNumIntDocConc As Long, dtDataConc As Date
Dim dTaxaFinanceiraConc As Double, iTipoDestinoConc As Integer, lFornCliDestinoConc As Long, iFilialDestinoConc As Integer

On Error GoTo Erro_GeracaoPedCompraConc_Le_Concorrencias

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 51225

    'Montar o SELECT dinamicamente
    Call GeracaoPedCompraConc_Le_Concorrencias1(objGeracaoPedidoCompra, sSelect)

    sDescricaoConc = String(STRING_DESCRICAO_CAMPO, 0)

    'Preparar parte fixa do SELECT
    lErro = GeracaoPedCompraConc_Le_Concorrencias2(lComando, sSelect, lCodigoConc, sDescricaoConc, lNumIntDocConc, dtDataConc, dTaxaFinanceiraConc, iTipoDestinoConc, lFornCliDestinoConc, iFilialDestinoConc)
    If lErro <> SUCESSO Then gError 51226

    'Complementa a passagem dos parâmetros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao dos bloqueios
    lErro = GeracaoPedCompraConc_Le_Concorrencias3(lComando, objGeracaoPedidoCompra)
    If lErro <> SUCESSO Then gError 51236

    'Processar os registros lidos
    lErro = GeracaoPedCompraConc_Le_Concorrencias4(lComando, objGeracaoPedidoCompra, lCodigoConc, sDescricaoConc, lNumIntDocConc, dtDataConc, dTaxaFinanceiraConc, iTipoDestinoConc, lFornCliDestinoConc, iFilialDestinoConc)
    If lErro <> SUCESSO And lErro <> 51248 Then gError 51247
    If lErro = 51248 Then gError 51249

    Call Comando_Fechar(lComando)

    GeracaoPedCompraConc_Le_Concorrencias = SUCESSO

    Exit Function

Erro_GeracaoPedCompraConc_Le_Concorrencias:

    GeracaoPedCompraConc_Le_Concorrencias = gErr

    Select Case gErr

        Case 51225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 51226, 51236, 51247
        
        Case 51249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_CONCORRENCIA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147592)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub GeracaoPedCompraConc_Le_Concorrencias1(ByVal objGeracaoPedidoCompra As ClassGeracaoPedCompraConc, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String

    'Se é para ler as Concorrências associadas a Requisição
    If objGeracaoPedidoCompra.iAssociadaRC = CONCORRENCIA_ASSOCIADA_RC Then
    
        sFields = "Concorrencia.Codigo, Concorrencia.Descricao, Concorrencia.NumIntDoc, Concorrencia.Data, Concorrencia.TaxaFinanceira, Concorrencia.TipoDestino, Concorrencia.FornCliDestino, Concorrencia.FilialDestino"
        sFrom = " FROM Concorrencia, Comprador, ItensConcorrencia, ItemRCItemConcorrencia, CotacaoItemConcorrencia, ItensCotacao "
        sWhere = " WHERE Concorrencia.NumIntDoc = ItensConcorrencia.Concorrencia AND ItensConcorrencia.NumIntDoc = ItemRCItemConcorrencia.ItemConcorrencia AND Concorrencia.Comprador = Comprador.Codigo AND Comprador.CodUsuario = ? AND ItensConcorrencia.NumIntDoc = CotacaoItemConcorrencia.ItemConcorrencia AND CotacaoItemConcorrencia.ItemCotacao = ItensCotacao.NumIntDoc"
        sOrderBy = " ORDER BY " & objGeracaoPedidoCompra.sOrdenacaoConc
        
    'Se não
    Else
    
        sFields = "Concorrencia.Codigo, Concorrencia.Descricao, Concorrencia.NumIntDoc, Concorrencia.Data, Concorrencia.TaxaFinanceira, Concorrencia.TipoDestino, Concorrencia.FornCliDestino, Concorrencia.FilialDestino"
        sFrom = " FROM Concorrencia, Comprador, ItensConcorrencia, ItensCotacao, CotacaoItemConcorrencia"
        sWhere = " WHERE Concorrencia.Comprador = Comprador.Codigo AND Comprador.CodUsuario = ? AND Concorrencia.NumIntDoc = ItensConcorrencia.Concorrencia AND ItensConcorrencia.NumIntDoc NOT IN (SELECT ItemConcorrencia FROM ItemRCItemConcorrencia) AND ItensConcorrencia.NumIntDoc = CotacaoItemConcorrencia.ItemConcorrencia AND CotacaoItemConcorrencia.ItemCotacao = ItensCotacao.NumIntDoc"
        sOrderBy = " ORDER BY " & objGeracaoPedidoCompra.sOrdenacaoConc
        
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND Concorrencia.FilialEmpresa = ?"

    If objGeracaoPedidoCompra.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND Concorrencia.Data >= ?"

    If objGeracaoPedidoCompra.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND Concorrencia.Data <= ?"

    If objGeracaoPedidoCompra.lCodigoDe <> 0 Then sWhere = sWhere & " AND Concorrencia.Codigo >= ?"

    If objGeracaoPedidoCompra.lCodigoAte <> 0 Then sWhere = sWhere & " AND Concorrencia.Codigo <= ?"

    sWhere = sWhere & " AND Concorrencia.TipoDestino = ?"

    If objGeracaoPedidoCompra.iSelecionaDestino <> 0 Then

        If objGeracaoPedidoCompra.lFornCliDestino <> 0 Then sWhere = sWhere & " AND Concorrencia.FornCliDestino = ? "
        If objGeracaoPedidoCompra.iFilialDestino <> 0 Then sWhere = sWhere & " AND Concorrencia.FilialDestino = ? "
    End If

    sSelect = "SELECT DISTINCT " & sFields & sFrom & sWhere

End Sub

Function GeracaoPedCompraConc_Le_Concorrencias2(ByVal lComando As Long, ByVal sSelect As String, vlCodigoConc As Variant, vsDescricaoConc As Variant, vlNumIntDocConc As Variant, vdtDataConc As Variant, vdTaxaFinanceiraConc As Variant, viTipoDestinoConc As Variant, vlFornCliDestinoConc As Variant, viFilialDestinoConc As Variant) As Long
'Inicia a execução do SELECT dinâmico

Dim lErro As Long
Dim ret As Integer

On Error GoTo Erro_GeracaoPedCompraConc_Le_Concorrencias2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 51227

    ret = Comando_BindVarInt(lComando, vlCodigoConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51228

    ret = Comando_BindVarInt(lComando, vsDescricaoConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51229

    ret = Comando_BindVarInt(lComando, vlNumIntDocConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51230

    ret = Comando_BindVarInt(lComando, vdtDataConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51231

    ret = Comando_BindVarInt(lComando, vdTaxaFinanceiraConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51232

    ret = Comando_BindVarInt(lComando, viTipoDestinoConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51233

    ret = Comando_BindVarInt(lComando, vlFornCliDestinoConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51234

    ret = Comando_BindVarInt(lComando, viFilialDestinoConc)
    If (ret <> AD_SQL_SUCESSO) Then gError 51235

    GeracaoPedCompraConc_Le_Concorrencias2 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraConc_Le_Concorrencias2:

    GeracaoPedCompraConc_Le_Concorrencias2 = gErr

    Select Case gErr

        Case 51227 To 51235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147593)

    End Select

    Exit Function

End Function

Function GeracaoPedCompraConc_Le_Concorrencias3(ByVal lComando As Long, ByVal objGeracaoPedCompra As ClassGeracaoPedCompraConc) As Long
'Trata variáveis a serem lidas no Banco de Dados

Dim lErro As Long
Dim vsComprador As Variant, viFilialEmpresa As Variant, vdtDataDe As Variant, vdtDataAte As Variant, vlCodigoDe As Variant, vlCodigoAte As Variant
Dim viTipoDestino As Variant, vlFornCliDestino As Variant, viFilialDestino As Variant

On Error GoTo Erro_GeracaoPedCompraConc_Le_Concorrencias3
            
    'Usuário
    If (giFilialEmpresa <> EMPRESA_TODA) Then
        vsComprador = gsUsuario
        lErro = Comando_BindVarInt(lComando, vsComprador)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51237
    End If

    'FilialEmpresa
    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51237
    End If
            
    'Data de
    If objGeracaoPedCompra.dtDataDe <> DATA_NULA Then
        vdtDataDe = objGeracaoPedCompra.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51238
    End If
    
    'Data Até
    If objGeracaoPedCompra.dtDataAte <> DATA_NULA Then
        vdtDataAte = objGeracaoPedCompra.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51239
    End If

    'Código de
    If objGeracaoPedCompra.lCodigoDe <> 0 Then
        vlCodigoDe = objGeracaoPedCompra.lCodigoDe
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51240
    End If

    'Código até
    If objGeracaoPedCompra.lCodigoAte <> 0 Then
        vlCodigoAte = objGeracaoPedCompra.lCodigoAte
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 51241
    End If

    viTipoDestino = objGeracaoPedCompra.iTipoDestino
    lErro = Comando_BindVarInt(lComando, viTipoDestino)
    If (lErro <> AD_SQL_SUCESSO) Then gError 51242

    'Se foi seleciondo um Tipo de Destino
    If objGeracaoPedCompra.iSelecionaDestino <> 0 Then

        If objGeracaoPedCompra.lFornCliDestino <> 0 Then
            vlFornCliDestino = objGeracaoPedCompra.lFornCliDestino
            lErro = Comando_BindVarInt(lComando, vlFornCliDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 51243
        End If

        If objGeracaoPedCompra.iFilialDestino <> 0 Then
            viFilialDestino = objGeracaoPedCompra.iFilialDestino
            lErro = Comando_BindVarInt(lComando, viFilialDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 51244
        End If
        
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 66777

    GeracaoPedCompraConc_Le_Concorrencias3 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraConc_Le_Concorrencias3:

    GeracaoPedCompraConc_Le_Concorrencias3 = gErr

    Select Case gErr

        Case 51237 To 51244, 66777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147594)

    End Select

    Exit Function

End Function

Function GeracaoPedCompraConc_Le_Concorrencias4(ByVal lComando As Long, ByVal objGeracaoPedCompra As ClassGeracaoPedCompraConc, lCodigoConc As Long, sDescricaoConc As String, lNumIntDocConc As Long, dtDataConc As Date, dTaxaFinanceiraConc As Double, iTipoDestinoConc As Integer, lFornCliDestinoConc As Long, iFilialDestinoConc As Integer) As Long
'Preenche a coleção de Concorrências de objGeracaoPedCompra com as concorrências lidas

Dim lErro As Long
Dim objConcorrencia As ClassConcorrencia

On Error GoTo Erro_GeracaoPedCompraConc_Le_Concorrencias4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 51245

    'Se não encontrou nenhuma Concorrência, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 51248

    'Enquanto encontrar concorrências
    Do While lErro = AD_SQL_SUCESSO

        Set objConcorrencia = New ClassConcorrencia

        objConcorrencia.lCodigo = lCodigoConc
        objConcorrencia.sDescricao = sDescricaoConc
        objConcorrencia.lNumIntDoc = lNumIntDocConc
        objConcorrencia.dtData = dtDataConc
        objConcorrencia.dTaxaFinanceira = dTaxaFinanceiraConc
        objConcorrencia.iTipoDestino = iTipoDestinoConc
        objConcorrencia.lFornCliDestino = lFornCliDestinoConc
        objConcorrencia.iFilialDestino = iFilialDestinoConc
        objConcorrencia.iFilialEmpresa = giFilialEmpresa

        'Adiciona na coleção de Concorrências
        objGeracaoPedCompra.colConcorrencia.Add objConcorrencia

        'Busca próxima concorrência
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 51246

    Loop

    GeracaoPedCompraConc_Le_Concorrencias4 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraConc_Le_Concorrencias4:

    GeracaoPedCompraConc_Le_Concorrencias4 = gErr

    Select Case gErr

        Case 51245, 51246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case 51248
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147595)

    End Select

    Exit Function

End Function

Function GerPedCompraConc_Le_Requisicoes(ByVal objGeracaoPedCompra As ClassGeracaoPedCompraConc, ByVal colRequisicoes As Collection) As Long
'Lê as Requisições associadas a concorrência

Dim lErro As Long
Dim lComando As Long
Dim objConcorrencia As ClassConcorrencia
Dim objItemConcorrencia As ClassItemConcorrencia
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objObservacao As New ClassObservacao
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_GerPedCompraConc_Le_Requisicoes

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 51250
                
    'Guarda concorrência selecionada
    Set objConcorrencia = objGeracaoPedCompra.colConcorrencia(objGeracaoPedCompra.iConcorrenciaSel)
                
    tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
    
    'Lê as Requisições associadas aos itens da concorrência passada por objConcorrencia
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RequisicaoCompra.NumIntDoc, Requisicaocompra.FilialEmpresa, Requisicaocompra.Codigo, RequisicaoCompra.Data, Requisicaocompra.DataLimite, RequisicaoCompra.Urgente, RequisicaoCompra.Requisitante, RequisicaoCompra.Ccl, RequisicaoCompra.TipoDestino, RequisicaoCompra.FornCliDestino, RequisicaoCompra.FilialDestino, RequisicaoCompra.Observacao FROM Requisicaocompra, ItemRCItemConcorrencia, ItensReqCompra, ItensConcorrencia, Concorrencia WHERE ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND ItensReqCompra.NumIntDoc = ItemRCItemConcorrencia.ItemReqCompra AND ItemRCItemConcorrencia.ItemConcorrencia = ItensConcorrencia.NumIntDoc AND ItensConcorrencia.Concorrencia = Concorrencia.NumIntDoc AND Concorrencia.NumIntDoc = ? ORDER BY RequisicaoCompra.Codigo, RequisicaoCompra.FilialEmpresa", _
    tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sCcl, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66789

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66767
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
                    
        Set objRequisicaoCompras = New ClassRequisicaoCompras
        
        objRequisicaoCompras.lCodigo = tRequisicaoCompras.lCodigo
        objRequisicaoCompras.iFilialEmpresa = tRequisicaoCompras.iFilialEmpresa
        objRequisicaoCompras.lNumIntDoc = tRequisicaoCompras.lNumIntDoc
        objRequisicaoCompras.dtData = tRequisicaoCompras.dtData
        objRequisicaoCompras.dtDataLimite = tRequisicaoCompras.dtDataLimite
        objRequisicaoCompras.lUrgente = tRequisicaoCompras.lUrgente
        objRequisicaoCompras.lRequisitante = tRequisicaoCompras.lRequisitante
        objRequisicaoCompras.sCcl = tRequisicaoCompras.sCcl
        objRequisicaoCompras.iTipoDestino = tRequisicaoCompras.iTipoDestino
        objRequisicaoCompras.lFornCliDestino = tRequisicaoCompras.lFornCliDestino
        objRequisicaoCompras.iFilialDestino = tRequisicaoCompras.iFilialDestino
        objRequisicaoCompras.iSelecionado = MARCADO
                    
        'Se a Requisição tem observação
        If tRequisicaoCompras.lObservacao <> 0 Then
              
            objObservacao.lNumInt = tRequisicaoCompras.lObservacao

            'Lê a observação
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 66935
            If lErro = 53827 Then gError 66936

            objRequisicaoCompras.sObservacao = objObservacao.sObservacao
            
        End If
        
        'Lê os itens de Requisição de Compras
        lErro = CF("ItensReqCompras_Le", objRequisicaoCompras)
        If lErro <> SUCESSO Then gError 66778
        
        'Adiciona nova Requisição de Compras
        colRequisicoes.Add objRequisicaoCompras
        
        'Busca próxima Requisição
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66768
    
    Loop
        
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    GerPedCompraConc_Le_Requisicoes = SUCESSO

    Exit Function

Erro_GerPedCompraConc_Le_Requisicoes:

    GerPedCompraConc_Le_Requisicoes = gErr

    Select Case gErr

        Case 51250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66778, 66935

        Case 66767, 66768, 66789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA2", gErr)
        
        Case 66936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, tRequisicaoCompras.lObservacao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147596)

    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensConcorrencia_Le(ByVal objConcorrencia As ClassConcorrencia) As Long
'Lê os itens da concorrencia passada como parâmetro

Dim lErro  As Long
Dim lComando As Long
Dim tItemConcorrencia As typeItemConcorrencia
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_ItensConcorrencia_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51252

    tItemConcorrencia.sProduto = String(STRING_PRODUTO, 0)
    tItemConcorrencia.sUM = String(STRING_UM_SIGLA, 0)
    tItemConcorrencia.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    'Seleciona os itens ligados a concorrência passada
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, Fornecedor, Filial, Quantidade, UM, DataNecessidade, Produtos.Descricao FROM ItensConcorrencia,Produtos WHERE ItensConcorrencia.Produto = Produtos.Codigo AND Concorrencia = ?", tItemConcorrencia.lNumIntDoc, tItemConcorrencia.sProduto, tItemConcorrencia.lFornecedor, tItemConcorrencia.iFilial, tItemConcorrencia.dQuantidade, tItemConcorrencia.sUM, tItemConcorrencia.dtDataNecessidade, tItemConcorrencia.sDescricao, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 51253

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51254

    'Para cada item lido
    Do While lErro = AD_SQL_SUCESSO

        'Preenche o objItemConcorrencia com os dados lidos
        Set objItemConcorrencia = New ClassItemConcorrencia

        objItemConcorrencia.lNumIntDoc = tItemConcorrencia.lNumIntDoc
        objItemConcorrencia.sProduto = tItemConcorrencia.sProduto
        objItemConcorrencia.lFornecedor = tItemConcorrencia.lFornecedor
        objItemConcorrencia.iFilial = tItemConcorrencia.iFilial
        objItemConcorrencia.dQuantidade = tItemConcorrencia.dQuantidade
        objItemConcorrencia.sUM = tItemConcorrencia.sUM
        objItemConcorrencia.dtDataNecessidade = tItemConcorrencia.dtDataNecessidade
        objItemConcorrencia.iEscolhido = MARCADO
        objItemConcorrencia.sDescricao = tItemConcorrencia.sDescricao

        'Carrega a coleção de CotacaoItemConcorrencia de objItemConcorrencia
        lErro = CF("CotacaoItemConcorrencia_Le", objItemConcorrencia)
        If lErro <> SUCESSO Then Error 51256

        'Lê os dadosdos itens de requisição que participaram
        'da geração desse item de concorrencia
        lErro = CF("ItemRCItemConcorrencia_Le", objItemConcorrencia)
        If lErro <> SUCESSO Then Error 51265

        'Carrega a coleção de quantidades suplementares do item da concorrência
        lErro = CF("QuantSuplementar_Le", objItemConcorrencia)
        If lErro <> SUCESSO Then Error 51266

        'Adiciona o objItemConcorrencia na coleção da Concorrência
        objConcorrencia.colItens.Add objItemConcorrencia

        'Pesquisa o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51255

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensConcorrencia_Le = SUCESSO

    Exit Function

Erro_ItensConcorrencia_Le:

    ItensConcorrencia_Le = gErr

    Select Case gErr

        Case 51252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", 0)

        Case 51253, 51254, 51255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 51256, 51265, 51266
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147597)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Usuarios_Le(objUsuarios As ClassUsuarios) As Long
'le  o usuário contido na tabela de Usuario pelo parametro CodUsuario e coloca no objeto
'Já existe na Tela de Pedido de Compras
End Function

Function PedidoCompra_LeMod(ByVal objPedidoCompra_BD As ClassPedidoCompras, ByVal lComando As Long) As Long
'Le o Pedido de Compra passado como parametro e locka,se encontrar

Dim lErro As Long
Dim tPedCompra As typePedidoCompras

On Error GoTo Erro_PedidoCompra_LeMod

    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)
    tPedCompra.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
    
    'Le o pedido de compra passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao,ValorProdutos,DataBaixa, Moeda, Taxa, UsuReg, UsuRegAprov, UsuRegEnvio " & _
                                        "FROM PedidoCompraN WHERE Codigo = ? AND FilialEmpresa=? AND Status = ?", 0, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dValorProdutos, tPedCompra.dtDataBaixa, tPedCompra.iMoeda, tPedCompra.dTaxa, tPedCompra.sUsuReg, tPedCompra.sUsuRegAprov, tPedCompra.sUsuRegEnvio, objPedidoCompra_BD.lCodigo, objPedidoCompra_BD.iFilialEmpresa, PEDIDOCOMPRA_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 62644

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62645

    'Se nao encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 62646

    'Preenche objPedidoCompra_BD
    objPedidoCompra_BD.lNumIntDoc = tPedCompra.lNumIntDoc
    objPedidoCompra_BD.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra_BD.dtData = tPedCompra.dtData
    objPedidoCompra_BD.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra_BD.lCodigo = tPedCompra.lCodigo
    objPedidoCompra_BD.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra_BD.iFilial = tPedCompra.iFilial
    objPedidoCompra_BD.iComprador = tPedCompra.iComprador
    objPedidoCompra_BD.sContato = tPedCompra.sContato
    objPedidoCompra_BD.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra_BD.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra_BD.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra_BD.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra_BD.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra_BD.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra_BD.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra_BD.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra_BD.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra_BD.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra_BD.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra_BD.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra_BD.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra_BD.lObservacao = tPedCompra.lObservacao
    objPedidoCompra_BD.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra_BD.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra_BD.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra_BD.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra_BD.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra_BD.sAlcada = tPedCompra.sAlcada
    objPedidoCompra_BD.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra_BD.dValorProdutos = tPedCompra.dValorProdutos
    objPedidoCompra_BD.iMoeda = tPedCompra.iMoeda
    objPedidoCompra_BD.dTaxa = tPedCompra.dTaxa
    objPedidoCompra_BD.sUsuReg = tPedCompra.sUsuReg
    objPedidoCompra_BD.sUsuRegAprov = tPedCompra.sUsuRegAprov
    objPedidoCompra_BD.sUsuRegEnvio = tPedCompra.sUsuRegEnvio
    
    'Lock do PedidoCompra
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 62647

    PedidoCompra_LeMod = SUCESSO

    Exit Function

Erro_PedidoCompra_LeMod:

    PedidoCompra_LeMod = Err

    Select Case Err

        Case 62644, 62645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra_BD.lCodigo)

        Case 62646

        Case 62647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOMPRA", Err, objPedidoCompra_BD.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147598)

    End Select

    Exit Function

End Function

Function Alcada_Le_Todas(ByVal colAlcada As Collection) As Long
'Guarda em colAlcada todas as alçadas cadastradas

Dim lErro As Long
Dim lComando As Long
Dim sCodigo As String
Dim objAlcada As ClassAlcada
Dim dLimiteOperacao As Double
Dim dLimiteMensal As Double

On Error GoTo Erro_Alcada_Le_Todas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49212

    sCodigo = String(STRING_ALCADA_CODUSUARIO, 0)

    'Seleciona todas as alcadas da tabela Alcada
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, LimiteMensal, LimiteOperacao FROM Alcada", sCodigo, dLimiteMensal, dLimiteOperacao)
    If lErro <> AD_SQL_SUCESSO Then Error 49213

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49214

    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm alçadas
        Set objAlcada = New ClassAlcada

        objAlcada.sCodUsuario = sCodigo
        objAlcada.dLimiteMensal = dLimiteMensal
        objAlcada.dLimiteOperacao = dLimiteOperacao

        colAlcada.Add objAlcada

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49216

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Alcada_Le_Todas = SUCESSO

    Exit Function

Erro_Alcada_Le_Todas:

    Alcada_Le_Todas = Err

    Select Case Err

        Case 49212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49213, 49214, 49216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALCADA", Err, objAlcada.sCodUsuario)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147599)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Alcada_Le(ByVal objAlcada As ClassAlcada) As Long
'Lê a alcada a partir do código do usuário.
'Devolve os dados em objAlcada.

Dim lErro As Long
Dim lComando As Long
Dim tAlcada As typeAlcada

On Error GoTo Erro_Alcada_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49205

    'Lê os dados da tabela Alcada
    lErro = Comando_Executar(lComando, "SELECT  LimiteOperacao,LimiteMensal FROM Alcada WHERE CodUsuario = ? ", tAlcada.dLimiteOperacao, tAlcada.dLimiteMensal, objAlcada.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 49206

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49207

    'Usuario nao possui alcada
    If lErro = AD_SQL_SEM_DADOS Then Error 49208

    'carrega os dados em objAlcada
    objAlcada.dLimiteMensal = tAlcada.dLimiteMensal
    objAlcada.dLimiteOperacao = tAlcada.dLimiteOperacao

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Alcada_Le = SUCESSO

    Exit Function

Erro_Alcada_Le:

    Alcada_Le = Err

    Select Case Err

        Case 49205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49206, 49207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49208
            'Erro tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147600)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicaoCompras_Le_Numero(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoCompras a Requisição passada por objRequisiçãoCompras com Código e FlialEmpresa

Dim lErro As Long
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long

On Error GoTo Erro_RequisicaoCompras_Le_Numero

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54436

    With tRequisicaoCompras

        .sCcl = String(STRING_CCL, 0)
        .sOPCodigo = String(STRING_OPCODIGO, 0)
        .sDigitador = String(STRING_USUARIO_CODIGO, 0)
        .sUsuReg = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
        
        'Lê os dados da Requisição da tabela RequisiçãoCompras
        '####################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, NumIntDocItemOP, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompra WHERE Codigo = ? AND FilialEmpresa = ?", _
            .lNumIntDoc, .iFilialEmpresa, .lCodigo, .dtData, .dtDataEnvio, .dtDataLimite, .lUrgente, .lRequisitante, .sDigitador, .sCcl, .iFilialCompra, .iTipoDestino, .lFornCliDestino, .iFilialDestino, .lObservacao, .iTipoTributacao, .sOPCodigo, .lNumIntDocItemOP, .sUsuReg, .sUsuRegAprov, .sUsuRegBaixa, .sUsuRegEnvio, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 54437
        '####################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54438

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 54439

        'Preenche o objRequisiçãoCompras
        objRequisicaoCompras.lNumIntDoc = .lNumIntDoc
        objRequisicaoCompras.iFilialEmpresa = .iFilialEmpresa
        objRequisicaoCompras.lCodigo = .lCodigo
        objRequisicaoCompras.dtData = .dtData
        objRequisicaoCompras.dtDataEnvio = .dtDataEnvio
        objRequisicaoCompras.dtDataLimite = .dtDataLimite
        objRequisicaoCompras.lUrgente = .lUrgente
        objRequisicaoCompras.lRequisitante = .lRequisitante
''''        objRequisicaoCompras.sDigitador = .sDigitador
        objRequisicaoCompras.sCcl = .sCcl
        objRequisicaoCompras.iFilialCompra = .iFilialCompra
        objRequisicaoCompras.iTipoDestino = .iTipoDestino
        objRequisicaoCompras.lFornCliDestino = .lFornCliDestino
        objRequisicaoCompras.iFilialDestino = .iFilialDestino
        objRequisicaoCompras.lObservacao = .lObservacao
        objRequisicaoCompras.sOPCodigo = .sOPCodigo
        objRequisicaoCompras.lNumIntDocItemOP = .lNumIntDocItemOP
        
        objRequisicaoCompras.sUsuReg = .sUsuReg
        objRequisicaoCompras.sUsuRegAprov = .sUsuRegAprov
        objRequisicaoCompras.sUsuRegBaixa = .sUsuRegBaixa
        objRequisicaoCompras.sUsuRegEnvio = .sUsuRegEnvio
        
    End With

    Call Comando_Fechar(lComando)

    RequisicaoCompras_Le_Numero = SUCESSO

    Exit Function

Erro_RequisicaoCompras_Le_Numero:

    RequisicaoCompras_Le_Numero = Err

    Select Case Err

        Case 54436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54437, 54438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54439 'Tratado na rotina de chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147601)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Usuario_Le_Todos(ByVal colUsuarios As Collection) As Long
'le todos os usuários contidos na tabela de Usuario e coloca os dados em colUsuarios

Dim lComando As Long
Dim lErro As Long
Dim tUsuario As typeUsuario
Dim objUsuario As ClassUsuario
    
On Error GoTo Erro_Usuario_Le_Todos

    lComando = 0

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 57293

    tUsuario.sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
    tUsuario.sNome = String(STRING_USUARIO, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    
    'Le os Usuarios ativos da tabela Usuario
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Nome, NomeReduzido FROM Usuario WHERE Ativo=?", tUsuario.sCodUsuario, tUsuario.sNome, tUsuario.sNomeReduzido, USUARIO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 57294
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57295
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objUsuario = New ClassUsuario
        
        'Passa para objUsuario os dados do Usuario lido da tabela
        objUsuario.sCodUsuario = tUsuario.sCodUsuario
        objUsuario.sNome = tUsuario.sNome
        objUsuario.sNomeReduzido = tUsuario.sNomeReduzido
        
        'Adiciona em colUsuarios o Usuario encontrado
        colUsuarios.Add objUsuario
        
        'Busca o proximo Usuario na tabela de Usuario
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57296
            
    Loop
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Usuario_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_Usuario_Le_Todos:

    Usuario_Le_Todos = Err

    Select Case Err
    
        Case 57293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 57294, 57295, 57296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147602)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Funcao alterada nesta tela (inclusao de campos na leitura da tabela)
Function Usuario_Le(ByVal objUsuario As ClassUsuario) As Long
'le os dados do usuario passado como parametro e devolve os dados em objUsuario

Dim lErro As Long
Dim lComando As Long
Dim tUsuario As typeUsuario

On Error GoTo Erro_Usuario_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36344
    
    tUsuario.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    
    'Le o usuario a partir do CodUsuario passado como parametro, da tabela de Usuario
    lErro = Comando_Executar(lComando, "SELECT Nome,NomeReduzido,Ativo FROM Usuario WHERE CodUsuario = ?", tUsuario.sNome, tUsuario.sNomeReduzido, tUsuario.iAtivo, objUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 36345
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36346
    
    'Se nao encontrou => erro
    If lErro <> AD_SQL_SUCESSO Then Error 36347
    
    'Passa os dados do Usuario para objUsuario
    objUsuario.sNome = tUsuario.sNome
    objUsuario.sNomeReduzido = tUsuario.sNomeReduzido
    objUsuario.iAtivo = tUsuario.iAtivo
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Usuario_Le = SUCESSO

    Exit Function

Erro_Usuario_Le:

    Usuario_Le = Err

    Select Case Err

        Case 36344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36345, 36346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
        
        Case 36347 'Usuario nao encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147603)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Usuario_Le_NomeRed(ByVal objUsuario As ClassUsuario) As Long
'Le o usuário a partir do nome reduzido passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim sNome As String
Dim sCodUsuario As String
Dim tUsuario As typeUsuario

On Error GoTo Erro_Usuario_Le_NomeRed

    'Abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 57266

    tUsuario.sCodUsuario = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    
    'Lê os dados da tabela Usuario
    lErro = Comando_Executar(lComando, "SELECT  CodUsuario,Nome,NomeReduzido,Ativo FROM Usuario WHERE NomeReduzido = ?", tUsuario.sCodUsuario, tUsuario.sNome, tUsuario.sNomeReduzido, tUsuario.iAtivo, objUsuario.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 57267

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57268

    'Usuario nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 57269

    'carrega os dados em objUsuario
    objUsuario.iAtivo = tUsuario.iAtivo
    objUsuario.sCodUsuario = tUsuario.sCodUsuario
    objUsuario.sNome = tUsuario.sNome
    objUsuario.sNomeReduzido = tUsuario.sNomeReduzido
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Usuario_Le_NomeRed = SUCESSO

    Exit Function

Erro_Usuario_Le_NomeRed:

    Usuario_Le_NomeRed = Err

    Select Case Err

        Case 57266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 57267, 57268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)

        Case 57269
            'Usuario nao cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147604)

    End Select

     Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BaixaPedCompras_ObterPedidos(ByVal objBaixaPedCompra As ClassBaixaPedCompra) As Long
'Preenche a colecao em objBaixaPedCompra com os pedidos que podem ser faturados e atendem
'aos criterios passados em objBaixaPedCompra com restricoes de No. de pedido, data de emissao,...

Dim lErro As Long, sSelect As String
Dim lComando As Long
Dim iFilialEmpresa As Integer
Dim iFilial As Integer
Dim iFilialForn As Integer
Dim iCodFilialForn As Integer
Dim lCodigo As Long
Dim lFornecedor As Long
Dim lCodForn As Long
Dim dtData As Date
Dim dtDataEnvio As Date
Dim dValorProdutos As Double
Dim sNomeReduzidoForn As String
Dim sNomeFilialForn As String
Dim lNumIntDoc As Long
Dim dtDataEnvio2 As Date
Dim sUsuario As String

On Error GoTo Erro_BaixaPedCompras_ObterPedidos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 57140

    'Montar SELECT dos pedidos de compra dinamicamente
    Call BaixaPedCompras_ObterPedidos1(objBaixaPedCompra, sSelect)

    sNomeReduzidoForn = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    sNomeFilialForn = String(STRING_FILIAL_FORNECEDOR_NOME, 0)
    dtData = DATA_NULA
    dtDataEnvio = DATA_NULA
    sUsuario = gsUsuario

    'Preparar parte fixa do SELECT
    lErro = BaixaPedCompras_ObterPedidos2(lComando, sSelect, iFilialEmpresa, lCodigo, lFornecedor, dtData, dtDataEnvio, iFilialForn, dValorProdutos, sNomeReduzidoForn, sNomeFilialForn, lNumIntDoc, dtDataEnvio2, sUsuario)
    If lErro <> SUCESSO Then gError 57141

    'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao dos bloqueios
    lErro = BaixaPedCompras_ObterPedidos3(lComando, objBaixaPedCompra)
    If lErro <> SUCESSO Then gError 57142

    'Processar os registros lidos
    lErro = BaixaPedCompras_ObterPedidos4(lComando, objBaixaPedCompra, iFilialEmpresa, lCodigo, lFornecedor, dtData, dtDataEnvio, iFilialForn, dValorProdutos, sNomeReduzidoForn, sNomeFilialForn, lNumIntDoc, dtDataEnvio2)
    If lErro <> SUCESSO And lErro <> 57166 Then gError 57143

    If lErro = 57166 Then gError 68394
    
    If objBaixaPedCompra.iSoResiduais = 1 Then
        
        'Lê os Itens de Pedidos de Compras que são Residuais e adiciona-os em objBaixaPedCompra.colPedCompras
        lErro = BaixaPedCompras_Obter_Residuais(objBaixaPedCompra)
        If lErro <> SUCESSO Then gError 63830
    
    End If
    
    'Fechar Comando
    lErro = Comando_Fechar(lComando)

    BaixaPedCompras_ObterPedidos = SUCESSO

    Exit Function

Erro_BaixaPedCompras_ObterPedidos:

    BaixaPedCompras_ObterPedidos = gErr

    Select Case gErr

        Case 57141 To 57143

        Case 63830
            'Erro tratado na rotina chamada
            
        Case 57140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_ENCONTRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147605)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub BaixaPedCompras_ObterPedidos1(ByVal objBaixaPedCompra As ClassBaixaPedCompra, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String, sGroupBy As String, sHaving As String

    sFields = "PedidoCompra.FilialEmpresa,PedidoCompra.Codigo, PedidoCompra.Fornecedor,PedidoCompra.Data, PedidoCompra.DataEnvio, PedidoCompra.Filial,PedidoCompra.ValorProdutos, Fornecedores.NomeReduzido, FiliaisFornecedores.Nome,PedidoCompra.NumIntDoc "
    sFrom = " FROM PedidoCompra, Fornecedores,FiliaisFornecedores, Compradores "
    sWhere = " WHERE  PedidoCompra.Fornecedor=FiliaisFornecedores.CodFornecedor AND PedidoCompra.Filial=FiliaisFornecedores.CodFilial AND PedidoCompra.Fornecedor = Fornecedores.Codigo AND Compradores.Codigo = PedidoCompra.Comprador AND PedidoCompra.DataEnvio <> ? AND Compradores.CodUsuario = ?"
    sOrderBy = " ORDER BY " & objBaixaPedCompra.sOrdenacao
    sHaving = ""

    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND PedidoCompra.FilialEmpresa = ?"

    If objBaixaPedCompra.lPedCompraDe <> 0 Then sWhere = sWhere & " AND PedidoCompra.Codigo >= ?"

    If objBaixaPedCompra.lPedCompraAte <> 0 Then sWhere = sWhere & " AND PedidoCompra.Codigo <= ?"

    If objBaixaPedCompra.lFornecedorDe <> 0 Then sWhere = sWhere & " AND PedidoCompra.Fornecedor >= ?"

    If objBaixaPedCompra.lFornecedorAte <> 0 Then sWhere = sWhere & " AND PedidoCompra.Fornecedor <= ?"

    If objBaixaPedCompra.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND PedidoCompra.Data >= ?"

    If objBaixaPedCompra.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND PedidoCompra.Data <= ?"

    If objBaixaPedCompra.dtDataEnvioDe <> DATA_NULA Then sWhere = sWhere & " AND PedidoCompra.DataEnvio >= ?"

    If objBaixaPedCompra.dtDataEnvioAte <> DATA_NULA Then sWhere = sWhere & " AND PedidoCompra.DataEnvio <= ?"

    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sGroupBy & sOrderBy

End Sub

 Function BaixaPedCompras_ObterPedidos2(ByVal lComando As Long, ByVal sSelect As String, vFilialEmpresa As Variant, vCodigo As Variant, vFornecedor As Variant, vData As Variant, vDataEnvio As Variant, vFilialForn As Variant, vValorProdutos As Variant, vNomeForn As Variant, vNomeFilialForn As Variant, vNumIntDoc As Variant, vDataEnvio2 As Variant, vsUsuario As Variant) As Long
'Preparar parte fixa do SELECT

Dim ret As Integer, lErro As Long

On Error GoTo Erro_BaixaPedCompras_ObterPedidos2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 57144

    ret = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then gError 57145

    ret = Comando_BindVarInt(lComando, vCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 57146

    ret = Comando_BindVarInt(lComando, vFornecedor)
    If (ret <> AD_SQL_SUCESSO) Then gError 57147

    ret = Comando_BindVarInt(lComando, vData)
    If (ret <> AD_SQL_SUCESSO) Then gError 57149

    ret = Comando_BindVarInt(lComando, vDataEnvio)
    If (ret <> AD_SQL_SUCESSO) Then gError 57150
   
    ret = Comando_BindVarInt(lComando, vFilialForn)
    If (ret <> AD_SQL_SUCESSO) Then gError 57148

    ret = Comando_BindVarInt(lComando, vValorProdutos)
    If (ret <> AD_SQL_SUCESSO) Then gError 57151

    ret = Comando_BindVarInt(lComando, vNomeForn)
    If (ret <> AD_SQL_SUCESSO) Then gError 57152

    ret = Comando_BindVarInt(lComando, vNomeFilialForn)
    If (ret <> AD_SQL_SUCESSO) Then gError 57153

    ret = Comando_BindVarInt(lComando, vNumIntDoc)
    If (ret <> AD_SQL_SUCESSO) Then gError 63001
    
    vDataEnvio2 = DATA_NULA
    ret = Comando_BindVarInt(lComando, vDataEnvio2)
    If (ret <> AD_SQL_SUCESSO) Then gError 68396
    
    ret = Comando_BindVarInt(lComando, vsUsuario)
    If (ret <> AD_SQL_SUCESSO) Then gError 68396

    BaixaPedCompras_ObterPedidos2 = SUCESSO

    Exit Function

Erro_BaixaPedCompras_ObterPedidos2:

    BaixaPedCompras_ObterPedidos2 = gErr

    Select Case gErr

        Case 57144 To 57153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr, vCodigo)

        Case 63001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr, vCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147606)

    End Select

    Exit Function

End Function

Function BaixaPedCompras_ObterPedidos3(ByVal lComando As Long, ByVal objBaixaPedCompra As ClassBaixaPedCompra) As Long
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao dos bloqueios

Dim lErro As Long
Dim viFilialEmpresa As Variant, vlPedidosDe As Variant, vlPedidosAte As Variant, vlFornecedorDe As Variant, vlFornecedorAte As Variant, vdtDataEnvioDe As Variant, vdtDataEnvioAte As Variant, vdtDataDe As Variant, vdtDataAte As Variant
Dim vdtDataEnvio2 As Variant

On Error GoTo Erro_BaixaPedCompras_ObterPedidos3

    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57155
    End If

    If objBaixaPedCompra.lPedCompraDe <> 0 Then
        vlPedidosDe = objBaixaPedCompra.lPedCompraDe
        lErro = Comando_BindVarInt(lComando, vlPedidosDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57156
    End If

    If objBaixaPedCompra.lPedCompraAte <> 0 Then
        vlPedidosAte = objBaixaPedCompra.lPedCompraAte
        lErro = Comando_BindVarInt(lComando, vlPedidosAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57157
    End If

    If objBaixaPedCompra.lFornecedorDe <> 0 Then
        vlFornecedorDe = objBaixaPedCompra.lFornecedorDe
        lErro = Comando_BindVarInt(lComando, vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57158
    End If

    If objBaixaPedCompra.lFornecedorAte <> 0 Then
        vlFornecedorAte = objBaixaPedCompra.lFornecedorAte
        lErro = Comando_BindVarInt(lComando, vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57159
    End If

    If objBaixaPedCompra.dtDataDe <> DATA_NULA Then
        vdtDataDe = objBaixaPedCompra.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57162
    End If

    If objBaixaPedCompra.dtDataAte <> DATA_NULA Then
        vdtDataAte = objBaixaPedCompra.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57163
    End If

    If objBaixaPedCompra.dtDataEnvioDe <> DATA_NULA Then
        vdtDataEnvioDe = objBaixaPedCompra.dtDataEnvioDe
        lErro = Comando_BindVarInt(lComando, vdtDataEnvioDe)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57160
    End If

    If objBaixaPedCompra.dtDataEnvioAte <> DATA_NULA Then
        vdtDataEnvioAte = objBaixaPedCompra.dtDataEnvioAte
        lErro = Comando_BindVarInt(lComando, vdtDataEnvioAte)
        If (lErro <> AD_SQL_SUCESSO) Then Error 57161
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 57164

    BaixaPedCompras_ObterPedidos3 = SUCESSO

    Exit Function

Erro_BaixaPedCompras_ObterPedidos3:

    BaixaPedCompras_ObterPedidos3 = Err

    Select Case Err

        Case 57155 To 57164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_COMPRA_BAIXA_PC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147607)

    End Select

    Exit Function

End Function

'??? dtDataEnvio2 é parametro nao utilizado: jones 2002
Function BaixaPedCompras_ObterPedidos4(ByVal lComando As Long, ByVal objBaixaPedCompra As ClassBaixaPedCompra, iFilialEmpresa As Integer, lCodigo As Long, lFornecedor As Long, dtData As Date, dtDataEnvio As Date, iFilial As Integer, dValorProdutos As Double, sNomeReduzidoForn As String, sNomeFilialForn As String, lNumIntDoc As Long, dtDataEnvio2 As Date) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objPedidoCompra As ClassPedidoCompras
Dim iContador As Integer

On Error GoTo Erro_BaixaPedCompras_ObterPedidos4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57165

    'Se não encontrou==>Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 57166

    iContador = 0

    Do While lErro = AD_SQL_SUCESSO

        iContador = iContador + 1

        'Verifica se o número de pedidos é superior ao número máximo de pedidos permitido
        'If iContador = NUM_MAX_PEDIDOS Then Error 57002

        Set objPedidoCompra = New ClassPedidoCompras

        objPedidoCompra.iFilialEmpresa = iFilialEmpresa
        objPedidoCompra.lCodigo = lCodigo
        objPedidoCompra.lFornecedor = lFornecedor
        objPedidoCompra.iFilial = iFilial
        objPedidoCompra.dtDataEnvio = dtDataEnvio
        objPedidoCompra.dtData = dtData
        objPedidoCompra.dValorProdutos = dValorProdutos
        objPedidoCompra.lNumIntDoc = lNumIntDoc

        'Adiciona objPedidoCompra na coleção de Pedidos de Compra
        objBaixaPedCompra.colPedCompras.Add objPedidoCompra

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57167

    Loop

    BaixaPedCompras_ObterPedidos4 = SUCESSO

    Exit Function

Erro_BaixaPedCompras_ObterPedidos4:

    BaixaPedCompras_ObterPedidos4 = Err

    Select Case Err

        Case 57165, 57167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra.lCodigo)

        Case 57166
            'Erro tratado na rotina chamadora

        Case 57002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDO_BAIXAR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147608)

    End Select

    Exit Function

End Function

Function BaixaPedCompras_Obter_Residuais(ByVal objBaixaPedCompra As ClassBaixaPedCompra) As Long
'Verifica se os itens de Pedido de Compra são Residuais e adiciona-os na colecaoPedCompras
' Para trazer as residuais:
' 1) olhar primeiro em produto se o resíduo está preenchido (<> -1)
' 2) Se em produto não estiver preenchido olhar se está em TipoDeproduto
' 3) Se não estiver preenchido em TP Produto olhar em Compras config
'* O Valor zero é um valor válido para cálculo de resíduo, logo,
'só qdo o resíduo estiver com -1 é que significa que ele não foi preenchido.

Dim lErro As Long
Dim objPedComprasInfo As New ClassPedidoCompras

Dim iIndice As Integer
Dim objItemPedCompras As New ClassItemPedCompra
Dim objPedidoCompras As New ClassPedidoCompras
Dim objPedidoCompras1 As New ClassPedidoCompras
Dim iResidual As Integer
Dim objProduto As New ClassProduto
Dim iIndice2 As Integer
Dim dMinPercRec As Double
Dim dResiduo As Double
Dim iNumItens As Integer

On Error GoTo Erro_BaixaPedCompras_Obter_Residuais

    'Percorre os elementos da colecao PedCompras
    For iIndice = objBaixaPedCompra.colPedCompras.Count To 1 Step -1

        Set objPedidoCompras = objBaixaPedCompra.colPedCompras(iIndice)

        objPedidoCompras1.lNumIntDoc = objPedidoCompras.lNumIntDoc
        
        Set objPedidoCompras1.colItens = New Collection
        
        'Lê os itens do Pedido de Compras
        lErro = CF("ItensPC_Le", objPedidoCompras1)
        If lErro <> SUCESSO Then Error 63833

        'Percorre os elementos de colItens, que contem os itens de um Pedido de Compra
        For iIndice2 = objPedidoCompras1.colItens.Count To 1 Step -1
            
            iNumItens = objPedidoCompras1.colItens.Count
            
            Set objItemPedCompras = objPedidoCompras1.colItens(iIndice2)

            objProduto.sCodigo = objItemPedCompras.sProduto

            'Verifica o Valor do Residuo para o Produto passado como parametro
            lErro = CF("Produto_Le_Residual", objProduto, dResiduo)
            If lErro <> SUCESSO Then Error 63834
         
            If objItemPedCompras.dQuantidade <> 0 Then dMinPercRec = objItemPedCompras.dQuantRecebida / objItemPedCompras.dQuantidade
                
           'Verifica se é um Item Residual
           If dResiduo + dMinPercRec >= 1 Then iResidual = iResidual + 1

        Next
        
        'Se o item não é residual
        If iResidual <> iNumItens Then

            'remove o Pedido de Compras da colecao
            objBaixaPedCompra.colPedCompras.Remove (iIndice)

        End If

        iResidual = 0
        iNumItens = 0
        
    Next

    BaixaPedCompras_Obter_Residuais = SUCESSO

    Exit Function

Erro_BaixaPedCompras_Obter_Residuais:

    BaixaPedCompras_Obter_Residuais = Err

    Select Case Err

        Case 63833, 63834, 63838
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147609)

        End Select

        Exit Function

End Function

Function PedidoCompra_LeLock(ByVal objPedidoCompra_BD As ClassPedidoCompras) As Long
'Le o Pedido de Compra passado como parametro e locka,se encontrar

Dim lErro As Long
Dim tPedCompra As typePedidoCompras
Dim lComando As Long

On Error GoTo Erro_PedidoCompra_LeLock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63617
    
    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)
    tPedCompra.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
    
    'Le o pedido de compra passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao, UsuReg, UsuRegAprov, UsuRegEnvio FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa=?", 0, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.sUsuReg, tPedCompra.sUsuRegAprov, tPedCompra.sUsuRegEnvio, objPedidoCompra_BD.lCodigo, objPedidoCompra_BD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 56257

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56258

    'Se nao encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 56259

    'Preenche objPedidoCompra_BD
    objPedidoCompra_BD.lNumIntDoc = tPedCompra.lNumIntDoc
    objPedidoCompra_BD.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra_BD.dtData = tPedCompra.dtData
    objPedidoCompra_BD.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra_BD.lCodigo = tPedCompra.lCodigo
    objPedidoCompra_BD.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra_BD.iFilial = tPedCompra.iFilial
    objPedidoCompra_BD.iComprador = tPedCompra.iComprador
    objPedidoCompra_BD.sContato = tPedCompra.sContato
    objPedidoCompra_BD.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra_BD.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra_BD.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra_BD.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra_BD.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra_BD.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra_BD.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra_BD.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra_BD.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra_BD.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra_BD.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra_BD.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra_BD.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra_BD.lObservacao = tPedCompra.lObservacao
    objPedidoCompra_BD.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra_BD.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra_BD.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra_BD.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra_BD.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra_BD.sAlcada = tPedCompra.sAlcada
    objPedidoCompra_BD.sUsuReg = tPedCompra.sUsuReg
    objPedidoCompra_BD.sUsuRegAprov = tPedCompra.sUsuRegAprov
    objPedidoCompra_BD.sUsuRegEnvio = tPedCompra.sUsuRegEnvio

    'Lock do PedidoCompra
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 56260

    Call Comando_Fechar(lComando)
    
    PedidoCompra_LeLock = SUCESSO

    Exit Function

Erro_PedidoCompra_LeLock:

    PedidoCompra_LeLock = Err

    Select Case Err

        Case 56257, 56258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra_BD.lCodigo)

        Case 56259

        Case 56260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOMPRA", Err, objPedidoCompra_BD.lCodigo)

        Case 63617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147610)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItensPC_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê no BD a coleção de ítens do Pedido de Compra e devolve em colItensPedido dentro de objPedidoCompra

Dim lErro As Long
Dim tItemPedido As typeItemPedCompra
Dim lComando As Long
Dim lTransacao As Long
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_ItensPC_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56107

    tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0) 'OK o tamanho correto da string é 255
    tItemPedido.sUM = String(STRING_UM_SIGLA, 0)

    'OK  comentário incorreto
    'Le da tabela de ItensPedCompra a colecao de itens do Pedido de Compra passado como parametro
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS,DeliveryDate, TempoTransito FROM ItensPedCompra WHERE PedCompra = ? ORDER BY NumIntDoc", tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, _
    tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, tItemPedido.dtDeliveryDate, tItemPedido.iTempoTransito, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 56109

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56110

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56111

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPC = New ClassItemPedCompra

        'preenchendo objitem
        objItemPC.dAliquotaICMS = tItemPedido.dAliquotaICMS
        objItemPC.dAliquotaIPI = tItemPedido.dAliquotaIPI
        objItemPC.dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        objItemPC.dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        objItemPC.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemPC.dQuantidade = tItemPedido.dQuantidade
        objItemPC.dQuantRecebida = tItemPedido.dQuantRecebida
        objItemPC.dQuantRecebimento = tItemPedido.dQuantRecebimento
        objItemPC.dtDataLimite = tItemPedido.dtDataLimite
        objItemPC.dValorDesconto = tItemPedido.dValorDesconto
        objItemPC.dValorIPI = tItemPedido.dValorIPI
        objItemPC.iRebebForaFaixa = tItemPedido.iRebebForaFaixa
        objItemPC.iStatus = tItemPedido.iStatus
        objItemPC.iTipoOrigem = tItemPedido.iTipoOrigem
        objItemPC.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemPC.lNumIntOrigem = tItemPedido.lNumIntOrigem
        objItemPC.lObservacao = tItemPedido.lObservacao
        objItemPC.sDescProduto = tItemPedido.sDescProduto
        objItemPC.sProduto = tItemPedido.sProduto
        objItemPC.sUM = tItemPedido.sUM
        
        objItemPC.dtDeliveryDate = tItemPedido.dtDeliveryDate
        objItemPC.iTempoTransito = tItemPedido.iTempoTransito
        
        lErro = CF("LocalizacaoItemPC_Le", objItemPC)
        If lErro <> SUCESSO And lErro <> 56361 Then Error 49490

        'Insere em colItens os valores de objItemPC
        objPedidoCompra.colItens.Add objItemPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56112
    Loop
    
    '######################################################
    'Inserido por Wagner 17/07/2006
    lErro = CF("PCItens_Le_Customizado", objPedidoCompra)
    If lErro <> SUCESSO Then gError 181224
    '######################################################

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPC_Le = SUCESSO

    Exit Function

Erro_ItensPC_Le:

    ItensPC_Le = gErr

    Select Case gErr

        Case 56107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56109, 56110, 56112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr, objItemPC.lNumIntDoc)

        Case 49490, 181224
            'Erro tratado na rotina chamadora

        'OK O erro 56111 não está sendo tratado aqui nem na rotina chamadora , trate-o aqui.
        Case 56111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objPedidoCompra.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147611)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function LocalizacaoItemPC_Le(ByVal objItemPC As ClassItemPedCompra) As Long
'Le ItemPC na tabela de LocalizacaoItemPC

Dim lErro As Long
Dim tItem As typeLocalizacaoItemPC
Dim lComando As Long
Dim lTransacao As Long
Dim objLocalizacao As ClassLocalizacaoItemPC

On Error GoTo Erro_LocalizacaoItemPC_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56357

    tItem.sCcl = String(STRING_CCL_DESCRICAO, 0)
    tItem.sContaContabil = String(STRING_CONTA, 0)

    'Busca na tabela de LocalizacaoItemPC  Localizacao dos Itens de Pedido de Compra passados como parametro
    lErro = Comando_Executar(lComando, "SELECT Almoxarifado,Ccl,Quantidade,ContaContabil FROM LocalizacaoItensPC WHERE ItemPC = ? ", tItem.iAlmoxarifado, tItem.sCcl, tItem.dQuantidade, tItem.sContaContabil, objItemPC.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 56359

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56360


    Do While lErro <> AD_SQL_SEM_DADOS

        Set objLocalizacao = New ClassLocalizacaoItemPC

        objLocalizacao.iAlmoxarifado = tItem.iAlmoxarifado
        objLocalizacao.dQuantidade = tItem.dQuantidade
        objLocalizacao.sCcl = tItem.sCcl
        objLocalizacao.sContaContabil = tItem.sContaContabil

        'Adiciona Localizacao em ItemPC
        objItemPC.colLocalizacao.Add objLocalizacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56362
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    LocalizacaoItemPC_Le = SUCESSO

    Exit Function

Erro_LocalizacaoItemPC_Le:

    LocalizacaoItemPC_Le = Err

    Select Case Err

        Case 56357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56359, 56360, 56362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", Err, objItemPC.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147612)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Residual(ByVal objProduto As ClassProduto, dResiduo As Double) As Long
'Verifica o Valor do Residuo para o Produto passado como parametro

' Para trazer as residuais:
' 1) olhar primeiro em produto se o resíduo está preenchido (<> -1)
' 2) Se em produto não estiver preenchido olhar se está em TipoDeproduto
' 3) Se não estiver preenchido em TP Produto olhar em Compras config
'* O Valor zero é um valor válido para cálculo de resíduo, logo,
'só qdo o resíduo estiver com -1 é que significa que ele não foi preenchido.

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Produto_Le_Residual

        'Lê o Produto do Item de Pedido de Compra
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 63834

        'Se não encontrou o Produto ==> erro
        If lErro = 28030 Then Error 63835

        'Verifica se o Residuo do Produto lido foi preenchido
        If objProduto.dResiduo <> -1 Then

            dResiduo = objProduto.dResiduo
        'Se o Residuo do Produto não foi preenchido
        ElseIf objProduto.dResiduo = -1 Then

            'Verifica se o Tipo do Produto foi definido
            If objProduto.iTipo > 0 Then
            
                objTipoDeProduto.iTipo = objProduto.iTipo

                'Lê o Tipo de Produto do Produto em questão
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then Error 63836

                'Se não encontrou o TipoDeProduto ==> erro
                If lErro = 22531 Then Error 63837

                'Verifica se o Residuo do TipoDeProduto lido foi preenchido
                If objTipoDeProduto.dResiduo <> -1 Then

                    dResiduo = objTipoDeProduto.dResiduo
                
                'Se o Residuo do TipoProduto não foi preenchido
                ElseIf objTipoDeProduto.dResiduo = -1 Then
            
                    'Preenche o Residuo com o valor default de Compras
                    dResiduo = gobjCOM.dResiduo
                End If
            
            ElseIf objProduto.iTipo = 0 Then
                
                'Preenche o Residuo com o valor default de Compras
                dResiduo = gobjCOM.dResiduo
            
            End If
        End If
    
    Produto_Le_Residual = SUCESSO
        
    Exit Function

Erro_Produto_Le_Residual:

    Produto_Le_Residual = Err
    
    Select Case Err
    
        Case 63834, 63836
            'Erros tratados nas rotinas chamadas
            
        Case 63835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 63837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", Err, objTipoDeProduto.iTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147613)
        
    End Select
    
    Exit Function
    
End Function

Function Valor_Produtos_Recebido(ByVal objPedidoCompra As ClassPedidoCompras, dValorProdutoRecebido As Double)
'Calcula o valor recebido efetivo dos produtos do Pedido de Compra passado como parâmetro

Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim dValorTotal As Double
Dim dQuantRecebidaEfetiva As Double
Dim dValorDescontoProporcional As Double
Dim dValorRecebidoEfetivo As Double

On Error GoTo Erro_Valor_Produtos_Recebido

    dValorProdutoRecebido = 0
    
    'Lê os itens do Pedido de Compra, cujo número do Pedido foi passado como parâmetro
    lErro = CF("ItensPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then Error 57493

    For Each objItemPC In objPedidoCompra.colItens

        'Calcula o valor total do Item do Pedido de Compra
        dValorTotal = (objItemPC.dQuantidade * objItemPC.dPrecoUnitario) - objItemPC.dValorDesconto

        'Calcula a quantidade efetiva recebida do item do Pedido de Compra
        dQuantRecebidaEfetiva = objItemPC.dQuantRecebida + objItemPC.dQuantRecebimento

        'Calcula o valor do desconto, proporcional à quantidade já recebida do item do Pedido de Compra,
        'O valor do desconto armazenado no banco de dados está referenciando o desconto total do produto,
        'não considerando a quantidade efetiva recebida
        If objItemPC.dQuantidade <> 0 Then dValorDescontoProporcional = (objItemPC.dValorDesconto * dQuantRecebidaEfetiva) / objItemPC.dQuantidade

        'Calcula o Valor Recebido Efetivo do item do Pedido de Compra
        dValorRecebidoEfetivo = (objItemPC.dPrecoUnitario * dQuantRecebidaEfetiva) - dValorDescontoProporcional

        'Calcula o Valor Recebido dos produtos do Pedido de Compra em questão
        dValorProdutoRecebido = dValorRecebidoEfetivo + dValorProdutoRecebido

    Next

    Exit Function

Erro_Valor_Produtos_Recebido:

    Valor_Produtos_Recebido = Err

    Select Case Err

        Case 57493
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147614)

    End Select

    Exit Function

End Function

Function BaixaPedCotacao_ObterPedidos(ByVal objBaixaPedCotacao As ClassBaixaPedCotacao) As Long
'Preenche a colecao em objPedCotacao com os pedidos que possuem Data de Emissão preenhida e
'possuem objetos associados (concorrência, Pedido de Compras) baixados,
'e atendem aos criterios passados na seleção

Dim lErro As Long, sSelect As String
Dim lComando As Long
Dim iFilialEmpresa As Integer
Dim iFilial As Integer
Dim iFilialForn As Integer
Dim iCodFilialForn As Integer
Dim lCodigo As Long
Dim lFornecedor As Long
Dim lCodForn As Long
Dim dtData As Date
Dim dtDataEmissao As Date
Dim dtDataValidade As Date
Dim lNumIntDoc As Long

On Error GoTo Erro_BaixaPedCotacao_ObterPedidos

    'Abetura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67570

    'Montar SELECT dos pedidos de Cotação dinamicamente
    Call BaixaPedCotacao_ObterPedidos1(objBaixaPedCotacao, sSelect)

    'Preparar parte fixa do SELECT
    lErro = BaixaPedCotacao_ObterPedidos2(lComando, sSelect, iFilialEmpresa, lCodigo, lFornecedor, dtData, dtDataEmissao, dtDataValidade, iFilialForn, lNumIntDoc)
    If lErro <> SUCESSO Then gError 67571

    'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao dos bloqueios
    lErro = BaixaPedCotacao_ObterPedidos3(lComando, objBaixaPedCotacao)
    If lErro <> SUCESSO Then gError 67572

    'Processar os registros lidos
    lErro = BaixaPedCotacao_ObterPedidos4(lComando, objBaixaPedCotacao, iFilialEmpresa, lCodigo, lFornecedor, dtData, dtDataEmissao, dtDataValidade, iFilialForn, lNumIntDoc)
    If lErro <> SUCESSO Then gError 67573
    
    'Fechar Comando
    Call Comando_Fechar(lComando)

    BaixaPedCotacao_ObterPedidos = SUCESSO

    Exit Function

Erro_BaixaPedCotacao_ObterPedidos:

    BaixaPedCotacao_ObterPedidos = gErr

    Select Case gErr

        Case 67571 To 67573
            
        Case 67570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147615)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub BaixaPedCotacao_ObterPedidos1(ByVal objBaixaPedCotacao As ClassBaixaPedCotacao, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String, sGroupBy As String, sHaving As String

    sFields = "PedidoCotacao.FilialEmpresa, PedidoCotacao.Codigo, PedidoCotacao.Fornecedor, PedidoCotacao.Data, PedidoCotacao.DataEmissao, PedidoCotacao.DataValidade, PedidoCotacao.Filial, PedidoCotacao.NumIntDoc"
    sFrom = " FROM PedidoCotacao"
    sWhere = " WHERE PedidoCotacao.NumIntDoc NOT IN (SELECT ItemPedCotacao.PedidoCotacao FROM ItemPedCotacao, ItensPedCompra, ItensCotacao WHERE ItensPedCompra.TipoOrigem = ? AND ItensPedCompra.NumIntOrigem = ItensCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc)"
    sOrderBy = " ORDER BY PedidoCotacao.Codigo"

    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND PedidoCotacao.FilialEmpresa = ?"

    If objBaixaPedCotacao.lPedCotacaoDe <> 0 Then sWhere = sWhere & " AND PedidoCotacao.Codigo >= ?"

    If objBaixaPedCotacao.lPedCotacaoAte <> 0 Then sWhere = sWhere & " AND PedidoCotacao.Codigo <= ?"

    If objBaixaPedCotacao.lFornecedorDe <> 0 Then sWhere = sWhere & " AND PedidoCotacao.Fornecedor >= ?"

    If objBaixaPedCotacao.lFornecedorAte <> 0 Then sWhere = sWhere & " AND PedidoCotacao.Fornecedor <= ?"

    If objBaixaPedCotacao.dtDataEmissaoDe <> DATA_NULA Then sWhere = sWhere & " AND PedidoCotacao.Data >= ?"

    If objBaixaPedCotacao.dtDataEmissaoAte <> DATA_NULA Then sWhere = sWhere & " AND PedidoCotacao.Data <= ?"

    If objBaixaPedCotacao.dtDataValidadeDe <> DATA_NULA Then sWhere = sWhere & " AND PedidoCotacao.DataValidade >= ?"

    If objBaixaPedCotacao.dtDataValidadeAte <> DATA_NULA Then sWhere = sWhere & " AND PedidoCotacao.DataValidade <= ?"

    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub

Function BaixaPedCotacao_ObterPedidos2(ByVal lComando As Long, ByVal sSelect As String, vFilialEmpresa As Variant, vCodigo As Variant, vFornecedor As Variant, vData As Variant, vDataEmissao As Variant, vDataValidade As Variant, vFilialForn As Variant, vNumIntDoc As Variant) As Long
'Preparar parte fixa do SELECT

Dim ret As Integer, lErro As Long

On Error GoTo Erro_BaixaPedCotacao_ObterPedidos2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 67574

    ret = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then gError 67575

    ret = Comando_BindVarInt(lComando, vCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 67576

    ret = Comando_BindVarInt(lComando, vFornecedor)
    If (ret <> AD_SQL_SUCESSO) Then gError 67577

    ret = Comando_BindVarInt(lComando, vData)
    If (ret <> AD_SQL_SUCESSO) Then gError 67578

    ret = Comando_BindVarInt(lComando, vDataEmissao)
    If (ret <> AD_SQL_SUCESSO) Then gError 67579
   
    ret = Comando_BindVarInt(lComando, vDataValidade)
    If (ret <> AD_SQL_SUCESSO) Then gError 67580
   
    ret = Comando_BindVarInt(lComando, vFilialForn)
    If (ret <> AD_SQL_SUCESSO) Then gError 67581

    ret = Comando_BindVarInt(lComando, vNumIntDoc)
    If (ret <> AD_SQL_SUCESSO) Then gError 67582
    
    BaixaPedCotacao_ObterPedidos2 = SUCESSO

    Exit Function

Erro_BaixaPedCotacao_ObterPedidos2:

    BaixaPedCotacao_ObterPedidos2 = gErr

    Select Case gErr

        Case 67574 To 67582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", gErr, vCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147616)

    End Select

    Exit Function

End Function

Function BaixaPedCotacao_ObterPedidos3(ByVal lComando As Long, ByVal objBaixaPedCotacao As ClassBaixaPedCotacao) As Long
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao dos bloqueios

Dim lErro As Long
Dim viFilialEmpresa As Variant, vlPedidosDe As Variant, vlPedidosAte As Variant, vlFornecedorDe As Variant, vlFornecedorAte As Variant, vDataValidadeDe As Variant, vDataValidadeAte As Variant, vDataEmissaoDe As Variant, vDataEmissaoAte As Variant
Dim vTipoOrigem As Variant, vDataEmissaoNula As Variant

On Error GoTo Erro_BaixaPedCotacao_ObterPedidos3

'    vDataEmissaoNula = DATA_NULA
'    lErro = Comando_BindVarInt(lComando, vDataEmissaoNula)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 67583
'
    vTipoOrigem = TIPO_ORIGEM_PEDCOTACAO
    lErro = Comando_BindVarInt(lComando, vTipoOrigem)
    If (lErro <> AD_SQL_SUCESSO) Then gError 67584
    
    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67585
    End If
    
    If objBaixaPedCotacao.lPedCotacaoDe <> 0 Then
        vlPedidosDe = objBaixaPedCotacao.lPedCotacaoDe
        lErro = Comando_BindVarInt(lComando, vlPedidosDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67586
    End If

    If objBaixaPedCotacao.lPedCotacaoAte <> 0 Then
        vlPedidosAte = objBaixaPedCotacao.lPedCotacaoAte
        lErro = Comando_BindVarInt(lComando, vlPedidosAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67587
    End If

    If objBaixaPedCotacao.lFornecedorDe <> 0 Then
        vlFornecedorDe = objBaixaPedCotacao.lFornecedorDe
        lErro = Comando_BindVarInt(lComando, vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67588
    End If

    If objBaixaPedCotacao.lFornecedorAte <> 0 Then
        vlFornecedorAte = objBaixaPedCotacao.lFornecedorAte
        lErro = Comando_BindVarInt(lComando, vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67589
    End If

    If objBaixaPedCotacao.dtDataEmissaoDe <> DATA_NULA Then
        vDataEmissaoDe = objBaixaPedCotacao.dtDataEmissaoDe
        lErro = Comando_BindVarInt(lComando, vDataEmissaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67590
    End If

    If objBaixaPedCotacao.dtDataEmissaoAte <> DATA_NULA Then
        vDataEmissaoAte = objBaixaPedCotacao.dtDataEmissaoAte
        lErro = Comando_BindVarInt(lComando, vDataEmissaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67591
    End If

    If objBaixaPedCotacao.dtDataValidadeDe <> DATA_NULA Then
        vDataValidadeDe = objBaixaPedCotacao.dtDataValidadeDe
        lErro = Comando_BindVarInt(lComando, vDataValidadeDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67592
    End If

    If objBaixaPedCotacao.dtDataValidadeAte <> DATA_NULA Then
        vDataValidadeAte = objBaixaPedCotacao.dtDataValidadeAte
        lErro = Comando_BindVarInt(lComando, vDataValidadeAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67593
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 67594

    BaixaPedCotacao_ObterPedidos3 = SUCESSO

    Exit Function

Erro_BaixaPedCotacao_ObterPedidos3:

    BaixaPedCotacao_ObterPedidos3 = gErr

    Select Case gErr

        Case 67583 To 67594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_COMPRA_BAIXA_PC", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147617)

    End Select

    Exit Function

End Function

Function BaixaPedCotacao_ObterPedidos4(ByVal lComando As Long, ByVal objBaixaPedCotacao As ClassBaixaPedCotacao, iFilialEmpresa As Integer, lCodigo As Long, lFornecedor As Long, dtData As Date, dtDataEmissao As Date, dtDataValidade As Date, iFilial As Integer, lNumIntDoc As Long) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objPedidoCotacao As ClassPedidoCotacao
Dim iContPedidos As Integer

On Error GoTo Erro_BaixaPedCotacao_ObterPedidos4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67595

    'Se não encontrou==>Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67596

    Do While lErro = AD_SQL_SUCESSO
        
        iContPedidos = iContPedidos + 1
        
        'Se o número de Pedidos ultrapassou o número máximo
        If iContPedidos = NUM_MAX_PEDCOTACOES Then gError 67598
        
        Set objPedidoCotacao = New ClassPedidoCotacao

        objPedidoCotacao.iFilialEmpresa = iFilialEmpresa
        objPedidoCotacao.lCodigo = lCodigo
        objPedidoCotacao.lFornecedor = lFornecedor
        objPedidoCotacao.iFilial = iFilial
        objPedidoCotacao.dtDataEmissao = dtDataEmissao
        objPedidoCotacao.dtData = dtData
        objPedidoCotacao.dtDataValidade = dtDataValidade
        objPedidoCotacao.lNumIntDoc = lNumIntDoc

        'Adiciona objPedidoCompra na coleção de Pedidos de Compra
        objBaixaPedCotacao.colPedCotacao.Add objPedidoCotacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67597

    Loop

    BaixaPedCotacao_ObterPedidos4 = SUCESSO

    Exit Function

Erro_BaixaPedCotacao_ObterPedidos4:

    BaixaPedCotacao_ObterPedidos4 = gErr

    Select Case gErr

        Case 67595, 67597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO1", gErr)

        Case 67596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDOCOTACAO", gErr)

        Case 67598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_PEDIDOS_SUPERIOR_MAXIMO", gErr, NUM_MAX_PEDCOTACOES)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147618)

    End Select

    Exit Function

End Function

'Function ItensReqCompras_Le(objReqCompras As ClassRequisicaoCompras) As Long
'
'Dim lErro As Long
'Dim tItemReqCompras As typeItemReqCompra
'Dim objItemReqCompra As ClassItemReqCompras
'Dim lComando As Long
'
'On Error GoTo Erro_ItensReqCompras_Le
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 51288
'
'    With tItemReqCompras
'
'        .sCcl = String(STRING_CCL, 0)
'        .sContaContabil = String(STRING_CONTA, 0)
'        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
'        .sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
'        .sProduto = String(STRING_PRODUTO, 0)
'        .sUM = String(STRING_UM_SIGLA, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Produto,DescProduto,Status,Quantidade,QuantPedida,QuantRecebida,QuantCancelada,UM,Ccl,Almoxarifado,ContaContabil,CreditaICMS,CreditaIPI,Observacao,Fornecedor,Filial,Exclusivo FROM ItensReqCompra WHERE ReqCompra = ?", .lNumIntDoc, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iCreditaICMS, .iCreditaIPI, .sObservacao, .lFornecedor, .iFilial, .iExclusivo)
'        If lErro <> AD_SQL_SUCESSO Then Error 51289
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51290
'
'        Do While lErro <> AD_SQL_SEM_DADOS
'
'            Set objItemReqCompra = New ClassItemReqCompras
'
'            objItemReqCompra.lNumIntDoc = .lNumIntDoc
'            objItemReqCompra.sProduto = .sProduto
'            objItemReqCompra.sDescProduto = .sDescProduto
'            objItemReqCompra.iStatus = .iStatus
'            objItemReqCompra.dQuantidade = .dQuantidade
'            objItemReqCompra.dQuantPedida = .dQuantPedida
'            objItemReqCompra.dQuantRecebida = .dQuantRecebida
'            objItemReqCompra.dQuantCancelada = .dQuantCancelada
'            objItemReqCompra.sUM = .sUM
'            objItemReqCompra.sCcl = .sCcl
'            objItemReqCompra.iAlmoxarifado = .iAlmoxarifado
'            objItemReqCompra.sContaContabil = .sContaContabil
'            objItemReqCompra.iCreditaICMS = .iCreditaICMS
'            objItemReqCompra.iCreditaIPI = .iCreditaIPI
'            objItemReqCompra.sObservacao = .sObservacao
'            objItemReqCompra.lFornecedor = .lFornecedor
'            objItemReqCompra.iFilial = .iFilial
'            objItemReqCompra.iExclusivo = .iExclusivo
'
'            objReqCompras.colItens.Add objItemReqCompra
'
'            lErro = Comando_BuscarProximo(lComando)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51291
'
'        Loop
'
'    End With
'
'    Call Comando_Fechar(lComando)
'
'    ItensReqCompras_Le = SUCESSO
'
'    Exit Function
'
'Erro_ItensReqCompras_Le:
'
'    Select Case Err
'
'        Case 51288
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 51289, 51290, 51291
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147619)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function Usuarios_Le_Nome(ByVal objUsuarios As ClassUsuarios) As Long
'le o usuário com o Nome passado e coloca no objUsuarios

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeUsuarios

On Error GoTo Erro_Usuarios_Le_Nome

    lComando = 0

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 50129

    tUsuarios.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuarios.sSenha = String(STRING_USUARIO_SENHA, 0)

    'Busca o usuario onde o nome esta no objUsuarios
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Senha, DataValidade, Ativo FROM Usuarios WHERE Nome=? And Ativo=?", tUsuarios.sCodUsuario, tUsuarios.sSenha, tUsuarios.dtDataValidade, tUsuarios.iAtivo, objUsuarios.sNome, Ativo)
    If lErro <> AD_SQL_SUCESSO Then Error 50130

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50131

    If lErro = AD_SQL_SEM_DADOS Then Error 50132

    'Coloca os dados lidos no objUsuarios
    objUsuarios.sCodUsuario = tUsuarios.sCodUsuario
    objUsuarios.sSenha = tUsuarios.sSenha
    objUsuarios.dtDataValidade = tUsuarios.dtDataValidade
    objUsuarios.iAtivo = tUsuarios.iAtivo

    Call Comando_Fechar(lComando)

    Usuarios_Le_Nome = SUCESSO

    Exit Function

Erro_Usuarios_Le_Nome:

    Usuarios_Le_Nome = Err

    Select Case Err

        Case 50129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50130, 50131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)

        Case 50132
            'Não encontrou usuário. Tratado na rotina chamadora.
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147620)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Usuarios_Le_NomeRed(ByVal objUsuarios As ClassUsuarios) As Long
'le o usuário com o Nome passado e coloca no objUsuarios

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeUsuarios

On Error GoTo Erro_Usuarios_Le_NomeRed

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 53203

    tUsuarios.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuarios.sSenha = String(STRING_USUARIO_SENHA, 0)
    tUsuarios.sNome = String(STRING_USUARIO_NOME, 0)

    'Busca o usuario onde o nome esta no objUsuarios
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Senha, DataValidade, Ativo, Nome FROM Usuarios WHERE NomeReduzido=? And Ativo=?", tUsuarios.sCodUsuario, tUsuarios.sSenha, tUsuarios.dtDataValidade, tUsuarios.iAtivo, tUsuarios.sNome, objUsuarios.sNomeReduzido, USUARIO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 53204

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53205

    If lErro = AD_SQL_SEM_DADOS Then Error 53206

    'Coloca os dados lidos no objUsuarios
    objUsuarios.sCodUsuario = tUsuarios.sCodUsuario
    objUsuarios.sSenha = tUsuarios.sSenha
    objUsuarios.dtDataValidade = tUsuarios.dtDataValidade
    objUsuarios.iAtivo = tUsuarios.iAtivo
    objUsuarios.sNome = tUsuarios.sNome
    
    Call Comando_Fechar(lComando)

    Usuarios_Le_NomeRed = SUCESSO

    Exit Function

Erro_Usuarios_Le_NomeRed:

    Usuarios_Le_NomeRed = Err

    Select Case Err

        Case 53203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53204, 53205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)

        Case 53206
            'Não encontrou usuário. Tratado na rotina chamadora.
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147621)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Comprador_Existe_Relacionamento(ByVal objComprador As ClassComprador) As Long
'Verifica se existe relacionamento com as tabelas Concorrencia,ConcorrenciaBaixada, PedidoCompra, PedidoCompraBaixado, Cotacao, CotacaoBaixada

Dim lErro As Long
Dim lComando As Long
Dim iCodComprador As Integer
Dim iStatus As Integer

On Error GoTo Erro_Comprador_Existe_Relacionamento

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 50108

    'Concorrencia
    lErro = Comando_Executar(lComando, "SELECT Comprador, Status FROM ConcorrenciaN WHERE Comprador= ?", iCodComprador, iStatus, objComprador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 50109

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50115

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then
    
        If iStatus = CONCORRENCIA_STATUS_ATIVO Then Error 50147
    
        If iStatus = CONCORRENCIA_STATUS_BAIXADO Then Error 50148

    End If

    'PedidoCompra
    lErro = Comando_Executar(lComando, "SELECT Comprador, Status FROM PedidoCompraN WHERE Comprador= ?", iCodComprador, iStatus, objComprador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 50111

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50117

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then
    
        If iStatus = PEDIDOCOMPRA_STATUS_ATIVO Then Error 50149
    
        If iStatus = PEDIDOCOMPRA_STATUS_BAIXADO Then Error 50150

    End If

    'Cotacao
    lErro = Comando_Executar(lComando, "SELECT Comprador, Status FROM CotacaoN WHERE Comprador= ?", iCodComprador, iStatus, objComprador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 50113

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50119

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then Error 50151

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then
    
        If iStatus = COTACAO_STATUS_ATIVO Then Error 50151
    
        If iStatus = COTACAO_STATUS_BAIXADO Then Error 50152

    End If

    Call Comando_Fechar(lComando)

    Comprador_Existe_Relacionamento = SUCESSO

    Exit Function

Erro_Comprador_Existe_Relacionamento:

    Comprador_Existe_Relacionamento = Err

    Select Case Err

        Case 50108, 50109, 50110, 50111, 50112, 50113, 50114

        Case 50115, 50116, 50117, 50118, 50119, 50120

        Case 50147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_CONCORRENCIA", Err, objComprador.iCodigo)

        Case 50148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_CONCORRENCIABAIXADA", Err, objComprador.iCodigo)

        Case 50149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_PEDIDOCOMPRA", Err, objComprador.iCodigo)

        Case 50150
                lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_PEDIDOCOMPRABAIXADO", Err, objComprador.iCodigo)

        Case 50151
                lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_COTACAO", Err, objComprador.iCodigo)

        Case 50152
                lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_VINCULADO_COTACAOBAIXADA", Err, objComprador.iCodigo)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147622)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Comprador_Le(ByVal objComprador As ClassComprador) As Long
'Le um comprador

Dim lErro As Long, iFilialEmpresa As Integer
Dim lComando As Long, iFilEmp As Integer
Dim sCodUsuario As String, sEmail As String
Dim iAumentaQuant As Integer

On Error GoTo Erro_Comprador_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 50061

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    sEmail = String(65, 0)
    
    iFilialEmpresa = objComprador.iFilialEmpresa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa

    'Seleciona os CodUsuario e AumentaQuant onde o Codigo for igual ao do objComprador e FilialEmpresa=giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, CodUsuario, AumentaQuant, Email FROM Comprador WHERE Codigo = ? and FilialEmpresa=?", iFilEmp, sCodUsuario, iAumentaQuant, sEmail, objComprador.iCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 50062

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50063

    If lErro = AD_SQL_SEM_DADOS Then Error 50064

    objComprador.iFilialEmpresa = iFilEmp
    objComprador.sCodUsuario = sCodUsuario
    objComprador.iAumentaQuant = iAumentaQuant
    objComprador.sEmail = sEmail

    Call Comando_Fechar(lComando)

    Comprador_Le = SUCESSO

    Exit Function

Erro_Comprador_Le:

    Comprador_Le = Err

    Select Case Err

        Case 50061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50062, 50063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRADOR", Err, objComprador.iCodigo)

        Case 50064 'Não existe na tabela de Compradores

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147623)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Comprador_Le_Todos(ByVal colComprador As Collection) As Long
'Lê todos os compradores da gifilialempresa e guarda em colComprador
'OK ** Mal descrito. FilialEmpresa ?

Dim lErro As Long
Dim lComando As Long
Dim sCodUsuario As String
Dim objComprador As ClassComprador
'Dim iCodigo As Double
Dim iAumentaQuant As Double
Dim iCodigo As Integer

On Error GoTo Erro_Comprador_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 50123

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)

    'OK ** Usuarios ?????
    'Seleciona todos oa compradores da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, AumentaQuant,CodUsuario FROM Comprador WHERE FilialEmpresa = ?", iCodigo, iAumentaQuant, sCodUsuario, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 50124

    lErro = Comando_BuscarPrimeiro(lComando)
    'OK ** Esse erro evita o seguinte
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50125
    If lErro = AD_SQL_SEM_DADOS Then Error 50126

    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm compradores
        Set objComprador = New ClassComprador

        objComprador.sCodUsuario = sCodUsuario
        objComprador.iCodigo = iCodigo
        objComprador.iAumentaQuant = iAumentaQuant

        colComprador.Add objComprador

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50127

    Loop

    Call Comando_Fechar(lComando)

    Comprador_Le_Todos = SUCESSO

    Exit Function

Erro_Comprador_Le_Todos:

    Comprador_Le_Todos = Err

    Select Case Err

        Case 50123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50124, 50125, 50127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRADOR2", Err)

        'OK ** Faltou "A ser tratado na rotina chamadora."
        Case 50126 'A ser tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147624)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Comprador_Le_Usuario(ByVal objComprador As ClassComprador) As Long
'Le o comprador de um usuario

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim iAumentaQuant As Integer
Dim sEmail As String
On Error GoTo Erro_Comprador_Le_Usuario

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 50056
    
    sEmail = String(65, 0)

    'Seleciona o Codigo do Comprador e se pode aumentar a quantidade do Usuario em objComprador
    lErro = Comando_Executar(lComando, "SELECT Codigo, AumentaQuant, Email FROM Comprador WHERE CodUsuario = ? AND FilialEmpresa = ?", iCodigo, iAumentaQuant, sEmail, objComprador.sCodUsuario, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 50057

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50058

    If lErro = AD_SQL_SEM_DADOS Then Error 50059

    objComprador.iAumentaQuant = iAumentaQuant
    objComprador.iCodigo = iCodigo
    objComprador.sEmail = sEmail

    Call Comando_Fechar(lComando)

    Comprador_Le_Usuario = SUCESSO

    Exit Function

Erro_Comprador_Le_Usuario:

    Comprador_Le_Usuario = Err

    Select Case Err

        Case 50056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50057, 50058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRADOR1", Err, objComprador.sCodUsuario)

        Case 50059 'Não existe na tabela de Compradores

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147625)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensConcorrenciaTodos_Le(ByVal objConcorrencia As ClassConcorrencia) As Long
'Lê os itens da concorrencia (baixada ou não) passada como parâmetro

Dim lErro  As Long, bConcorrenciaBaixada As Boolean
Dim lComando As Long
Dim tItemConcorrencia As typeItemConcorrencia
Dim objItemConcorrencia As ClassItemConcorrencia

On Error GoTo Erro_ItensConcorrenciaTodos_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67792

    tItemConcorrencia.sProduto = String(STRING_PRODUTO, 0)
    tItemConcorrencia.sUM = String(STRING_UM_SIGLA, 0)
    tItemConcorrencia.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    'Seleciona os itens ligados a concorrência passada
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, Fornecedor, Filial, Quantidade, UM, DataNecessidade,Produtos.Descricao FROM ItensConcorrenciaTodos,Produtos WHERE Concorrencia = ? AND Produtos.Codigo = ItensConcorrenciaTodos.Produto", tItemConcorrencia.lNumIntDoc, tItemConcorrencia.sProduto, tItemConcorrencia.lFornecedor, tItemConcorrencia.iFilial, tItemConcorrencia.dQuantidade, tItemConcorrencia.sUM, tItemConcorrencia.dtDataNecessidade, tItemConcorrencia.sDescricao, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67798

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67793

    'Para cada item lido
    Do While lErro = AD_SQL_SUCESSO

        'Preenche o objItemConcorrencia com os dados lidos
        Set objItemConcorrencia = New ClassItemConcorrencia

        objItemConcorrencia.lNumIntDoc = tItemConcorrencia.lNumIntDoc
        objItemConcorrencia.sProduto = tItemConcorrencia.sProduto
        objItemConcorrencia.lFornecedor = tItemConcorrencia.lFornecedor
        objItemConcorrencia.iFilial = tItemConcorrencia.iFilial
        objItemConcorrencia.dQuantidade = tItemConcorrencia.dQuantidade
        objItemConcorrencia.sUM = tItemConcorrencia.sUM
        objItemConcorrencia.dtDataNecessidade = tItemConcorrencia.dtDataNecessidade
        objItemConcorrencia.sDescricao = tItemConcorrencia.sDescricao
        
        'Carrega a coleção de CotacaoItemConcorrencia de objItemConcorrencia
        lErro = CotacaoItemConcorrenciaTodos_Le(objItemConcorrencia)
        If lErro <> SUCESSO Then gError 67794

        'Lê os dadosdos itens de requisição que participaram
        'da geração desse item de concorrencia
        lErro = CF("ItemRCItemConcorrencia_Le", objItemConcorrencia)
        If lErro <> SUCESSO Then gError 67795

        'Carrega a coleção de quantidades suplementares do item da concorrência
        lErro = CF("QuantSuplementar_Le", objItemConcorrencia, bConcorrenciaBaixada)
        If lErro <> SUCESSO Then gError 67796

        'Adiciona o objItemConcorrencia na coleção da Concorrência
        objConcorrencia.colItens.Add objItemConcorrencia

        'Pesquisa o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67797

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensConcorrenciaTodos_Le = SUCESSO

    Exit Function

Erro_ItensConcorrenciaTodos_Le:

    ItensConcorrenciaTodos_Le = gErr

    Select Case gErr

        Case 67793, 67797, 67798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCONCORRENCIATODOS", gErr)

        Case 67794, 67795, 67796
        
        Case 67792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", 0)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147626)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidosCotacaoTodos_Le_Concorrencia(ByVal objConcorrencia As ClassConcorrencia, ByVal colItensCotacao As Collection) As Long
'Lê os itens de cotação (baixados ou não) vinculadas a concorrência

Dim lErro As Long
Dim lComando As Long
Dim objCotItemConcAux As ClassCotacaoItemConcAux
Dim tCotItemConcAux As typeCotacaoItemConcAux
Dim objCondPagto As New ClassCondicaoPagto
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objMotivo As New AdmCodigoNome
Dim objItemConcorrencia  As ClassItemConcorrencia

On Error GoTo Erro_PedidosCotacaoTodos_Le_Concorrencia

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67799

    tCotItemConcAux.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tCotItemConcAux.sProduto = String(STRING_PRODUTO, 0)
    tCotItemConcAux.sUM = String(STRING_UM_SIGLA, 0)
    tCotItemConcAux.sMotivoEscolha = String(STRING_MOTIVO_ESCOLHA, 0)

    'Pesquisa no Banco, as cotações mais recentes vinculadas a concorrências
    lErro = Comando_Executar(lComando, "SELECT CotacaoProdutoTodas.Produto, MIN(Produtos.Descricao), ItensCotacaoTodos.CondPagto, MIN(ItensConcorrenciaTodos.Quantidade), MAX(CotacaoProdutoTodas.UM), MAX(ItensCotacaoTodos.PrecoUnitario), MAX(CotacaoItemConcorrencia.PrecoAjustado), MAX(CotacaoItemConcorrencia.ValorPresente), PedidoCotacaoTodos.Fornecedor, PedidoCotacaoTodos.Filial, MAX(PedidoCotacaoTodos.Codigo), MAX(PedidoCotacaoTodos.DataValidade), MAX(ItensCotacaoTodos.PrazoEntrega), MAX(ItensCotacaoTodos.QuantEntrega), MAX(ItensConcorrenciaTodos.DataNecessidade), MIN(CotacaoItemConcorrencia.MotivoEscolha), MAX(CotacaoTodas.Data), MIN(CotacaoItemConcorrencia.Escolhido), MIN(CotacaoProdutoTodas.Quantidade), MAX(ItensCotacaoTodos.NumIntDoc), MIN(CotacaoItemConcorrencia.DataEntrega), ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProdutoTodas, Produtos, ItensCotacaoTodos, PedidoCotacaoTodos, CotacaoItemConcorrencia, ItensConcorrenciaTodos, ConcorrenciaTodas, ItemPedCotacaoTodos, CotacaoTodas " & _
    " WHERE CotacaoProdutoTodas.Cotacao = CotacaoTodas.NumIntDoc AND CotacaoProdutoTodas.Produto = Produtos.Codigo AND CotacaoProdutoTodas.Produto = ItensConcorrenciaTodos.Produto AND CotacaoProdutoTodas.NumIntDoc = ItemPedCotacaoTodos.CotacaoProduto AND ItensConcorrenciaTodos.Concorrencia = ConcorrenciaTodas.NumIntDoc AND ItensConcorrenciaTodos.NumIntDoc = CotacaoItemConcorrencia.ItemConcorrencia AND CotacaoItemConcorrencia.ItemCotacao = ItensCotacaoTodos.NumIntDoc AND ItensCotacaoTodos.ItemPedCotacao = ItemPedCotacaoTodos.NumIntDoc AND ItemPedCotacaoTodos.PedidoCotacao = PedidoCotacaoTodos.NumIntDoc AND ConcorrenciaTodas.NumIntDoc = ? Group By CotacaoProdutoTodas.Produto, ItensCotacaoTodos.CondPagto,PedidoCotacaoTodos.Fornecedor, PedidoCotacaoTodos.Filial", _
    tCotItemConcAux.sProduto, tCotItemConcAux.sDescProduto, tCotItemConcAux.iCondPagto, tCotItemConcAux.dQuantComprar, tCotItemConcAux.sUM, tCotItemConcAux.dPrecoUnitario, tCotItemConcAux.dPrecoAjustado, tCotItemConcAux.dValorPresente, tCotItemConcAux.lFornecedor, tCotItemConcAux.iFilialForn, tCotItemConcAux.lPedidoCot, tCotItemConcAux.dtDataValidade, tCotItemConcAux.iPrazoEntrega, _
    tCotItemConcAux.dQuantidadeEntrega, tCotItemConcAux.dtDataNecessidade, tCotItemConcAux.sMotivoEscolha, tCotItemConcAux.dtDataCotacao, tCotItemConcAux.iEscolhido, tCotItemConcAux.dQuantComprarMax, tCotItemConcAux.lItemCotacao, tCotItemConcAux.dtDataEntrega, tCotItemConcAux.iMoeda, tCotItemConcAux.dTaxa, objConcorrencia.lNumIntDoc)
    If lErro <> SUCESSO Then gError 67800

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67801

    'Equanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objCotItemConcAux = New ClassCotacaoItemConcAux
        Set objCotItemConcAux.objCotacaoItemConc = New ClassCotacaoItemConc
        objItemConcorrencia = New ClassItemConcorrencia

        'Preenche objCotItemConcAux com os valores lidos
        objItemConcorrencia.sProduto = tCotItemConcAux.sProduto
        objItemConcorrencia.sDescricao = tCotItemConcAux.sDescProduto
        
        'Lê a Condição de pagamento
        objCondPagto.iCodigo = tCotItemConcAux.iCondPagto
        lErro = CF("CondicaoPagto_Le", objCondPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 67802

        'se nao encontrou --->erro
        If lErro = 19205 Then gError 67803

        objCotItemConcAux.objCotacaoItemConc.sCondPagto = objCondPagto.sDescReduzida
        objItemConcorrencia.sUM = tCotItemConcAux.sUM
        objCotItemConcAux.objCotacaoItemConc.dPrecoUnitario = tCotItemConcAux.dPrecoUnitario
        objCotItemConcAux.objCotacaoItemConc.dPrecoAjustado = tCotItemConcAux.dPrecoAjustado
        objCotItemConcAux.objCotacaoItemConc.dValorPresente = tCotItemConcAux.dValorPresente
        objCotItemConcAux.objCotacaoItemConc.iEscolhido = tCotItemConcAux.iEscolhido
        objCotItemConcAux.objCotacaoItemConc.lItemCotacao = tCotItemConcAux.lItemCotacao
        objCotItemConcAux.objCotacaoItemConc.iMoeda = tCotItemConcAux.iMoeda
        objCotItemConcAux.objCotacaoItemConc.dTaxa = tCotItemConcAux.dTaxa

        If tCotItemConcAux.lFornecedor <> 0 Then

            'Lê o Fornecedor
            objFornecedor.lCodigo = tCotItemConcAux.lFornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 67804

            'Se não encontrou, Erro
            If lErro = 12729 Then gError 67805

            objCotItemConcAux.objCotacaoItemConc.sFornecedor = objFornecedor.sNomeReduzido

        End If

        If tCotItemConcAux.iFilialForn <> 0 Then

            'Lê a Filial do Fornecedor
            objFilialFornecedor.iCodFilial = tCotItemConcAux.iFilialForn
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 67806

            'Se não encontrou, Erro
            If lErro = 12929 Then gError 67807

            objCotItemConcAux.objCotacaoItemConc.sFilial = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        End If

        objItemConcorrencia.dQuantidade = tCotItemConcAux.dQuantComprar
        objCotItemConcAux.objCotacaoItemConc.dQuantidadeComprar = tCotItemConcAux.dQuantComprarMax
        objCotItemConcAux.objCotacaoItemConc.lPedCotacao = tCotItemConcAux.lPedidoCot
        objCotItemConcAux.objCotacaoItemConc.dtDataValidade = tCotItemConcAux.dtDataValidade
        objCotItemConcAux.objCotacaoItemConc.iPrazoEntrega = tCotItemConcAux.iPrazoEntrega
        objCotItemConcAux.objCotacaoItemConc.dQuantEntrega = tCotItemConcAux.dQuantidadeEntrega
        objItemConcorrencia.dtDataNecessidade = tCotItemConcAux.dtDataNecessidade
        objCotItemConcAux.objCotacaoItemConc.dtDataEntrega = tCotItemConcAux.dtDataEntrega

        If Len(Trim(tCotItemConcAux.sMotivoEscolha)) > 0 Then

            'Lê o Motivo de Escolha
            objMotivo.sNome = tCotItemConcAux.sMotivoEscolha
            lErro = CF("Motivo_Le_Descricao", objMotivo)
            If lErro <> SUCESSO And lErro <> 66967 Then gError 67808

            'Se não encontrou, Erro
            If lErro = 66967 Then gError 67810

            objCotItemConcAux.objCotacaoItemConc.sMotivoEscolha = objMotivo.iCodigo & SEPARADOR & objMotivo.sNome

        End If

        'Guarda campos que serão usados na fução de ordenação de coleções
        If tCotItemConcAux.iFilialForn <> 0 Then objCotItemConcAux.sFornecedor = objFornecedor.sNomeReduzido
        If tCotItemConcAux.lFornecedor <> 0 Then objCotItemConcAux.sFilial = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        objCotItemConcAux.sCondPagto = objCondPagto.sDescReduzida
        objCotItemConcAux.sProduto = tCotItemConcAux.sProduto

        'Adiciona na coleção
        colItensCotacao.Add objCotItemConcAux

        'Busca próxima cotação
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67809

    Loop

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    PedidosCotacaoTodos_Le_Concorrencia = SUCESSO

    Exit Function

Erro_PedidosCotacaoTodos_Le_Concorrencia:

    PedidosCotacaoTodos_Le_Concorrencia = gErr

    Select Case gErr

        Case 67799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67800, 67801, 67809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOTODAS", gErr)

        Case 67802, 67804, 67806, 67808

        Case 67803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondPagto.iCodigo)

        Case 67805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 67807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 67810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_CADASTRADO1", gErr, objMotivo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147627)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoTributacaoPadrao_Le(iTipoTrib As Integer) As Long
'Lê os Tipo de Tributação Padrão a partir das Tabelas NaOpPadrao e FiliaisEmpresa
'e devolve em iTipoTrib

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TipoTributacaoPadrao_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66586

    'Lê Tipo de Tributação Padrão
    lErro = Comando_Executar(lComando, "SELECT NatOpPadrao.TipoTribEmp FROM NatOpPadrao, FiliaisEmpresa WHERE NatOpPadrao.TipoOperacao = ? AND NatOpPadrao.TipoAtividadeEmp = FiliaisEmpresa.TipoTribAtividade AND NatOpPadrao.Padrao = 1", iTipoTrib, TIPO_COMPRAS)
    If lErro <> AD_SQL_SUCESSO Then gError 66595

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66596

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66597

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TipoTributacaoPadrao_Le = SUCESSO

    Exit Function

Erro_TipoTributacaoPadrao_Le:

    TipoTributacaoPadrao_Le = gErr

    Select Case gErr

        Case 66586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66595, 66596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATOPPADRAO", gErr)

        Case 66597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147628)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposTributacao_Le(ByVal iTipo As Integer, sDescricao As String) As Long
'Lê Tipos de Tributação a partir do Tipo passado em iTipo

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TiposTributacao_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54351

    sDescricao = String(STRING_DESCRICAO_TIPOTRIBUTACAO, 0)

    'Lê os Tipos de Tributação associados a Compras
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM TiposDeTributacaoMovto WHERE Tipo = ?", sDescricao, iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 54351

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54352

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 54271

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposTributacao_Le = SUCESSO

    Exit Function

Erro_TiposTributacao_Le:

    TiposTributacao_Le = gErr

    Select Case gErr

        Case 54271 'Erro tratado na rotina chamada
        
        Case 54351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 54351, 54352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147629)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Observacao_Le(ByVal objObservacao As ClassObservacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim sObservacao As String

On Error GoTo Erro_Observacao_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 53824

    sObservacao = String(STRING_OBSERVACAO, 0)

    'Le na tabela de Observacao a Observacao cujo NumInt foi passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Observacao FROM Observacao WHERE NumInt=?", sObservacao, objObservacao.lNumInt)
    If lErro <> AD_SQL_SUCESSO Then Error 53825

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53826

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 53827

    objObservacao.sObservacao = sObservacao

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Observacao_Le = SUCESSO

    Exit Function

Erro_Observacao_Le:

    Observacao_Le = gErr

    Select Case gErr

        Case 53824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 53825, 53826
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBSERVACAO", gErr, objPedidoCompra.lObservacao)

        Case 53827
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_INEXISTENTE", gErr, objPedidoCompra.lObservacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147630)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensReqComprasTodos_Le(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras (baixada ou não) passada por objRequisiçãoCompras

Dim lErro As Long
Dim tItemReqCompras As typeItemReqCompra
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long
Dim iItem As Integer

On Error GoTo Erro_ItensReqComprasTodos_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67787

    With tItemReqCompras

        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)

        'Lê o ítem da Requisição na tabela ItensReqComprasTodos
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompraTodos WHERE ReqCompra = ? ORDER BY NumIntDoc", _
            .lNumIntDoc, .lReqCompra, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iTipoTributacao, .lObservacao, .lFornecedor, .iFilial, .iExclusivo, objRequisicaoCompras.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 67788

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67789

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67790
        
        iItem = 0
        'Equanto encontrar
        Do While lErro = AD_SQL_SUCESSO
        
            Set objItemReqCompras = New ClassItemReqCompras
            
            iItem = iItem + 1
            'Preenche o objItemReqCompras
            objItemReqCompras.lNumIntDoc = .lNumIntDoc
            objItemReqCompras.lReqCompra = .lReqCompra
            objItemReqCompras.sProduto = .sProduto
            objItemReqCompras.sDescProduto = .sDescProduto
            objItemReqCompras.iStatus = .iStatus
            objItemReqCompras.dQuantidade = .dQuantidade
            objItemReqCompras.dQuantPedida = .dQuantPedida
            objItemReqCompras.dQuantRecebida = .dQuantRecebida
            objItemReqCompras.dQuantCancelada = .dQuantCancelada
            objItemReqCompras.sUM = .sUM
            objItemReqCompras.sCcl = .sCcl
            objItemReqCompras.iAlmoxarifado = .iAlmoxarifado
            objItemReqCompras.sContaContabil = .sContaContabil
            objItemReqCompras.iTipoTributacao = .iTipoTributacao
            objItemReqCompras.lObservacao = .lObservacao
            objItemReqCompras.lFornecedor = .lFornecedor
            objItemReqCompras.iFilial = .iFilial
            objItemReqCompras.iExclusivo = .iExclusivo
            objItemReqCompras.iItem = iItem

            'Adiciona o objItemReqCompras na coleção colItens do objRequisiçãoCompras
            objRequisicaoCompras.colItens.Add objItemReqCompras

            'Busca no BD o próximo Ítem de Requisição de compras
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67791

        Loop

    End With

    Call Comando_Fechar(lComando)

    ItensReqComprasTodos_Le = SUCESSO

    Exit Function

Erro_ItensReqComprasTodos_Le:

    ItensReqComprasTodos_Le = gErr

    Select Case gErr

        Case 67790

        Case 67787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67788, 67789, 67791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRATODOS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147631)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CondicaoPagto_Le_DescReduzida(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
'le a Condicao de Pagto a partir da descrição reduzida

Dim lComando As Long
Dim lErro As Long
Dim tCondicaoPagto As typeCondicaoPagto

On Error GoTo Erro_CondicaoPagto_Le_DescReduzida

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67001

    tCondicaoPagto.sDescricao = String(STRING_CONDICAO_PAGTO_DESCRICAO, 0)
    
     'Verifica se a CondicaoPagto existe
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, EmPagamento, EmRecebimento, NumeroParcelas, DiasParaPrimeiraParcela, IntervaloParcelas, Mensal, DiaDoMes, AcrescimoFinanceiro FROM CondicoesPagto WHERE DescReduzida = ?", tCondicaoPagto.iCodigo, tCondicaoPagto.sDescricao, tCondicaoPagto.iEmPagamento, tCondicaoPagto.iEmRecebimento, tCondicaoPagto.iNumeroParcelas, tCondicaoPagto.iDiasParaPrimeiraParcela, tCondicaoPagto.iIntervaloParcelas, tCondicaoPagto.iMensal, tCondicaoPagto.iDiaDoMes, tCondicaoPagto.dAcrescimoFinanceiro, objCondicaoPagto.sDescReduzida)
    If lErro <> AD_SQL_SUCESSO Then gError 67002

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67003

    'CondicaoPagto não está cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 67004

    objCondicaoPagto.iCodigo = tCondicaoPagto.iCodigo
    objCondicaoPagto.sDescricao = tCondicaoPagto.sDescricao
    objCondicaoPagto.iEmPagamento = tCondicaoPagto.iEmPagamento
    objCondicaoPagto.iEmRecebimento = tCondicaoPagto.iEmRecebimento
    objCondicaoPagto.iNumeroParcelas = tCondicaoPagto.iNumeroParcelas
    objCondicaoPagto.iDiasParaPrimeiraParcela = tCondicaoPagto.iDiasParaPrimeiraParcela
    objCondicaoPagto.iIntervaloParcelas = tCondicaoPagto.iIntervaloParcelas
    objCondicaoPagto.iMensal = tCondicaoPagto.iMensal
    objCondicaoPagto.iDiaDoMes = tCondicaoPagto.iDiaDoMes
    objCondicaoPagto.dAcrescimoFinanceiro = tCondicaoPagto.dAcrescimoFinanceiro
    
    lErro = Comando_Fechar(lComando)

    CondicaoPagto_Le_DescReduzida = SUCESSO

    Exit Function

Erro_CondicaoPagto_Le_DescReduzida:

    CondicaoPagto_Le_DescReduzida = gErr

    Select Case gErr

        Case 67022
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67023, 67024
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO", gErr, objCondicaoPagto.iCodigo)

        Case 67025 'CondicaoPagto nao cadastrada (sem dados)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147632)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoItemConcorrencia_Le(ByVal objItemConcorrencia As ClassItemConcorrencia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tCotacaoItemConc As typeCotacaoItemConcorrencia
Dim objCotacaoItemConc As ClassCotacaoItemConc
Dim sSelect As String
Dim iFilial As Integer
Dim iCondPagto As Integer
Dim dQuantCotada As Double
Dim dtDataCotacao As Date
Dim sUM As String

On Error GoTo Erro_CotacaoItemConcorrencia_Le

'??? Jones: mesmo sabendo se a concorrencia está baixada ou nao, como vou saber se devo ligar à versao baixada ou nao das seguintes tabelas:
'??? ItensCotacao, ItemPedCotacao, PedidoCotacao, CotacaoProduto, Cotacao

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51257

    'Inserido UM - Wagner
    sSelect = "SELECT CotacaoItemConcorrencia.ItemCotacao,CotacaoItemConcorrencia.NumIntDoc,CotacaoItemConcorrencia.ValorPresente,CotacaoItemConcorrencia.Escolhido,CotacaoItemConcorrencia.MotivoEscolha,CotacaoItemConcorrencia.QuantidadeComprar,CotacaoItemConcorrencia.DataEntrega,CotacaoItemConcorrencia.PrecoAjustado,ItensCotacao.CondPagto, ItensCotacao.AliquotaIPI,ItensCotacao.AliquotaICMS, Fornecedores.NomeReduzido,FiliaisFornecedores.CodFilial, FiliaisFornecedores.Nome,PedidoCotacao.Codigo, CotacaoProduto.Quantidade, CondicoesPagto.DescReduzida,PedidoCotacao.DataValidade,Cotacao.Data,ItensCotacao.PrazoEntrega,ItensCotacao.QuantEntrega,ItensCotacao.PrecoUnitario, ItensCotacao.Moeda, ItensCotacao.Taxa, CotacaoProduto.UM " & _
              " FROM CotacaoItemConcorrencia, ItensCotacao, ItemPedCotacao, PedidoCotacao, Fornecedores, FiliaisFornecedores, CotacaoProduto, CondicoesPagto, Cotacao " & _
              " WHERE ItensCotacao.NumIntDoc = CotacaoItemConcorrencia.ItemCotacao AND ItemPedCotacao.NumIntDoc = ItensCotacao.ItemPedCotacao AND PedidoCotacao.NumIntDoc = ItemPedCotacao.PedidoCotacao AND PedidoCotacao.Fornecedor = fornecedores.codigo AND FiliaisFornecedores.CodFilial = PedidoCotacao.Filial AND FiliaisFornecedores.codfornecedor = Fornecedores.Codigo AND Cotacaoproduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND CondicoesPagto.Codigo = ItensCotacao.CondPagto AND Cotacao.numIntDoc = CotacaoProduto.Cotacao AND CotacaoItemConcorrencia.Itemconcorrencia = ?"
        
    tCotacaoItemConc.sMotivoEscolha = String(STRING_MOTIVO_ESCOLHA, 0)
    tCotacaoItemConc.sFilial = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCotacaoItemConc.sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    tCotacaoItemConc.sCondPagto = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
    sUM = String(STRING_UM_SIGLA, 0) 'Inserido por Wagner
    
    'Inserido UM - Wagner
    'Busca em CotacaoItemConcorrencia os registros ligados ao item passado
    lErro = Comando_Executar(lComando, sSelect, tCotacaoItemConc.lItemCotacao, tCotacaoItemConc.lNumIntDoc, tCotacaoItemConc.dValorPresente, tCotacaoItemConc.iEscolhido, tCotacaoItemConc.sMotivoEscolha, tCotacaoItemConc.dQuantidadeComprar, tCotacaoItemConc.dtDataEntrega, tCotacaoItemConc.dPrecoAjustado, iCondPagto, tCotacaoItemConc.dAliquotaIPI, tCotacaoItemConc.dAliquotaICMS, tCotacaoItemConc.sFornecedor, iFilial, tCotacaoItemConc.sFilial, tCotacaoItemConc.lPedCotacao, dQuantCotada, tCotacaoItemConc.sCondPagto, tCotacaoItemConc.dtDataValidade, dtDataCotacao, tCotacaoItemConc.iPrazoEntrega, tCotacaoItemConc.dQuantEntrega, tCotacaoItemConc.dPrecoUnitario, tCotacaoItemConc.iMoeda, tCotacaoItemConc.dTaxa, sUM, objItemConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 51258

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51259

    'Para cada registro encontrado
    Do While lErro <> AD_SQL_SEM_DADOS

        'Preenche um objCotacaoItemConc com os dados lidos
        Set objCotacaoItemConc = New ClassCotacaoItemConc

        With tCotacaoItemConc

            objCotacaoItemConc.dtDataValidade = .dtDataValidade
            objCotacaoItemConc.dtDataPedidoCotacao = dtDataCotacao
            objCotacaoItemConc.dCreditoICMS = .dCreditoICMS
            objCotacaoItemConc.dCreditoIPI = .dCreditoIPI
            objCotacaoItemConc.dPrecoAjustado = .dPrecoAjustado
            objCotacaoItemConc.dPrecoUnitario = .dPrecoUnitario
            objCotacaoItemConc.dPreferencia = .dPreferencia
            objCotacaoItemConc.dQuantEntrega = .dQuantEntrega
            objCotacaoItemConc.dQuantidadeComprar = .dQuantidadeComprar
            objCotacaoItemConc.dtDataEntrega = .dtDataEntrega
            objCotacaoItemConc.dtDataValidade = .dtDataValidade
            objCotacaoItemConc.dValorPresente = .dValorPresente
            objCotacaoItemConc.iEscolhido = .iEscolhido
            objCotacaoItemConc.iSelecionada = MARCADO
            objCotacaoItemConc.iPrazoEntrega = .iPrazoEntrega
            objCotacaoItemConc.lNumIntDoc = .lNumIntDoc
            objCotacaoItemConc.lItemCotacao = .lItemCotacao
            objCotacaoItemConc.sMotivoEscolha = .sMotivoEscolha
            objCotacaoItemConc.dAliquotaICMS = .dAliquotaICMS
            objCotacaoItemConc.dAliquotaIPI = .dAliquotaIPI
            objCotacaoItemConc.sCondPagto = iCondPagto & SEPARADOR & tCotacaoItemConc.sCondPagto '??? ler s
            objCotacaoItemConc.sFilial = iFilial & SEPARADOR & tCotacaoItemConc.sFilial
            objCotacaoItemConc.sFornecedor = .sFornecedor
            objCotacaoItemConc.lPedCotacao = .lPedCotacao
            objCotacaoItemConc.dQuantCotada = dQuantCotada
            objCotacaoItemConc.iMoeda = .iMoeda
            objCotacaoItemConc.dTaxa = .dTaxa
            objCotacaoItemConc.sUMCompra = sUM 'WAGNER
            
        End With

        'Adiciona o objCotacaoItemConc na coleção do item de concorrência
        objItemConcorrencia.colCotacaoItemConc.Add objCotacaoItemConc

        'Busca o proximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51260

    Loop

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    CotacaoItemConcorrencia_Le = SUCESSO

    Exit Function

Erro_CotacaoItemConcorrencia_Le:

    CotacaoItemConcorrencia_Le = gErr

    Select Case gErr

        Case 51257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 51258, 51259, 51260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147633)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemRCItemConcorrencia_Le(ByVal objItemConcorrencia As ClassItemConcorrencia) As Long
'Busca em ItemRCItemConcorrencia o link entre os itens de concorrência e os itens de requisição

Dim lErro As Long
Dim lComando As Long
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim dQuantidade As Double
Dim lItemReqCompra As Long
Dim lItemConcorrencia As Long

On Error GoTo Erro_ItemRCItemConcorrencia_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51261

    'Seleciona os registros vinculados ao item de concorrência
    lErro = Comando_Executar(lComando, "SELECT ItemReqCompra, ItemConcorrencia, Quantidade FROM ItemRCItemConcorrencia WHERE ItemConcorrencia = ?", lItemReqCompra, lItemConcorrencia, dQuantidade, objItemConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 51262

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51263

    'Para cada item lido
    Do While lErro <> AD_SQL_SEM_DADOS

        'Preenche o objItemRCItemConc com os dados lidos do BD
        Set objItemRCItemConc = New ClassItemRCItemConcorrencia

        objItemRCItemConc.dQuantidade = dQuantidade
        objItemRCItemConc.lItemReqCompra = lItemReqCompra

        'Adiciona o item na coleção do item de concorrência
        objItemConcorrencia.colItemRCItemConcorrencia.Add objItemRCItemConc

        'Busca pelo próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51264

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemRCItemConcorrencia_Le = SUCESSO

    Exit Function

Erro_ItemRCItemConcorrencia_Le:

    ItemRCItemConcorrencia_Le = gErr

    Select Case gErr

        Case 51261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 51262, 51263, 51264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147634)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function QuantSuplementar_Le(ByVal objItemConcorrencia As ClassItemConcorrencia, Optional ByVal bConcorrenciaBaixada As Boolean = False) As Long
'Lê as quantidades Suplementares associadas ao Item de Concorrência passado

Dim lErro As Long
Dim objQuantSuplementar As ClassQuantSuplementar
Dim tQuantSup As typeQuantSuplementar
Dim lComando As Long

On Error GoTo Erro_QuantSuplementar_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51267

    'Seleciona as quantidades suplementares do item passado
    lErro = Comando_Executar(lComando, "SELECT TipoDestino,FornCliDestino,FilialDestino,Quantidade FROM " & IIf(bConcorrenciaBaixada, "QuantidadesSuplementaresBaixa", "QuantidadesSuplementares") & " WHERE ItemConcorrencia = ?", tQuantSup.iTipoDestino, tQuantSup.lFornCliDestino, tQuantSup.iFilialDestino, tQuantSup.dQuantidade, objItemConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 51268

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51269

    Do While lErro <> AD_SQL_SEM_DADOS

        'Preenche o objQuantSuplementar com os dados lidos
        Set objQuantSuplementar = New ClassQuantSuplementar

        objQuantSuplementar.dQuantidade = tQuantSup.dQuantidade
        objQuantSuplementar.iFilialDestino = tQuantSup.iFilialDestino
        objQuantSuplementar.iTipoDestino = tQuantSup.iTipoDestino
        objQuantSuplementar.lFornCliDestino = tQuantSup.lFornCliDestino

        'adiciona o objQuantSuplementar na coleção do item de concorrência
        objItemConcorrencia.colQuantSuplementar.Add objQuantSuplementar

        'Pesquisa o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 51270

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    QuantSuplementar_Le = SUCESSO

    Exit Function

Erro_QuantSuplementar_Le:

    QuantSuplementar_Le = gErr

    Select Case gErr

        Case 51267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 51268, 51269, 51270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_QUANTIDADESSUPLEMENTARES", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147635)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Motivo_Le_Descricao(ByVal objMotivo As AdmCodigoNome) As Long
'Lê o Motivo a partir da descrição passada na tabela Motivo

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Motivo_Le_Descricao

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66964

    'Lê o código do Motivo da tabela Motivo
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Motivo", iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 66965

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66966

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66967

    objMotivo.iCodigo = iCodigo

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Motivo_Le_Descricao = SUCESSO

    Exit Function

Erro_Motivo_Le_Descricao:

    Motivo_Le_Descricao = gErr

    Select Case gErr

        Case 66964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66965, 66966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOTIVO", gErr)

        Case 66967 'Motivo não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147636)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicoesTodas_Le_Concorrencia(ByVal objConcorrencia As ClassConcorrencia, ByVal colRequisicoes As Collection) As Long
'Lê as Requisições (baixadas ou não) vinculadas aos Itens Concorrência  da Concorrência passada e devolve elas em colRequisicoes

Dim lErro As Long
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objRequisicaoCompras As ClassRequisicaoCompras
Dim lComando As Long
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_RequisicoesTodas_Le_Concorrencia

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67725
    
    tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
        
    'Lê as Requisições associadas ao item concorrência
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RequisicaoCompraTodas.NumIntDoc, RequisicaoCompraTodas.FilialEmpresa, RequisicaoCompraTodas.Codigo, RequisicaoCompraTodas.Data, RequisicaoCompraTodas.DataLimite, RequisicaoCompraTodas.Urgente, RequisicaoCompraTodas.Requisitante, RequisicaoCompraTodas.Ccl, RequisicaoCompraTodas.TipoDestino, RequisicaoCompraTodas.FornCliDestino, RequisicaoCompraTodas.FilialDestino, RequisicaoCompraTodas.Observacao FROM RequisicaoCompraTodas, ItemRCItemConcorrencia, ItensReqCompraTodos, ItensConcorrenciaTodos, ConcorrenciaTodas WHERE ItensReqCompraTodos.ReqCompra = RequisicaoCompraTodas.NumIntDoc AND ItensReqCompraTodos.NumIntDoc = ItemRCItemConcorrencia.ItemReqCompra AND ItemRCItemConcorrencia.ItemConcorrencia = ItensConcorrenciaTodos.NumIntDoc AND ItensConcorrenciaTodos.Concorrencia = ConcorrenciaTodas.NumIntDoc AND ConcorrenciaTodas.NumIntDoc = ? ORDER BY RequisicaoCompraTodas.Codigo, RequisicaoCompraTodas.FilialEmpresa", _
    tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sCcl, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67726
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67727
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
                    
        Set objRequisicaoCompras = New ClassRequisicaoCompras
        
        objRequisicaoCompras.lCodigo = tRequisicaoCompras.lCodigo
        objRequisicaoCompras.iFilialEmpresa = tRequisicaoCompras.iFilialEmpresa
        objRequisicaoCompras.lNumIntDoc = tRequisicaoCompras.lNumIntDoc
        objRequisicaoCompras.dtData = tRequisicaoCompras.dtData
        objRequisicaoCompras.dtDataLimite = tRequisicaoCompras.dtDataLimite
        objRequisicaoCompras.lUrgente = tRequisicaoCompras.lUrgente
        objRequisicaoCompras.lRequisitante = tRequisicaoCompras.lRequisitante
        objRequisicaoCompras.sCcl = tRequisicaoCompras.sCcl
        objRequisicaoCompras.iTipoDestino = tRequisicaoCompras.iTipoDestino
        objRequisicaoCompras.lFornCliDestino = tRequisicaoCompras.lFornCliDestino
        objRequisicaoCompras.iFilialDestino = tRequisicaoCompras.iFilialDestino
                    
        'Se a Requisição tem observação
        If tRequisicaoCompras.lObservacao <> 0 Then
              
            objObservacao.lNumInt = tRequisicaoCompras.lObservacao

            'Lê a observação
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 67730
            If lErro = 53827 Then gError 67729

            objRequisicaoCompras.sObservacao = objObservacao.sObservacao
            
        End If
        
        'Lê os itens de Requisição de Compras
        lErro = CF("ItensReqComprasTodos_Le", objRequisicaoCompras)
        If lErro <> SUCESSO Then gError 67731
        
        'Adiciona nova Requisição de Compras
        colRequisicoes.Add objRequisicaoCompras
        
        'Busca próxima Requisição
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67728
    
    Loop
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    RequisicoesTodas_Le_Concorrencia = SUCESSO
    
    Exit Function
    
Erro_RequisicoesTodas_Le_Concorrencia:

    RequisicoesTodas_Le_Concorrencia = gErr
    
    Select Case gErr
    
        Case 67725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 67726, 67727, 67728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA2", gErr)

        Case 67729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, tRequisicaoCompras.lObservacao)

        Case 67730, 67731
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147637)
        
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PedidoCompraTodos_Le_Concorrencia(ByVal objConcorrencia As ClassConcorrencia, ByVal colPedidos As Collection) As Long
'Lê os Pedidos vinculados ao ItemConcorrência da Concorrência passada e devolve eles em colPedidos

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim sUM As String
Dim dQuantidade As Double
Dim sProduto As String
Dim iMoeda As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo

On Error GoTo Erro_PedidoCompraTodos_Le_Concorrencia

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67732
                
    sUM = String(STRING_UM_SIGLA, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    'Lê os Pedidos de Compra associados as Itens Da Concorrência
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompraTodos.Moeda, PedidoCompraTodos.Codigo, ItensPedCompraTodos.Produto, ItensPedCompraTodos.Quantidade, ItensPedCompraTodos.UM FROM PedidoCompraTodos, ItensPedCompraTodos, CotacaoItemConcorrenciaTodos, ItensConcorrenciaTodos, ConcorrenciaTodas WHERE ItensPedCompraTodos.PedCompra = PedidoCompraTodos.NumIntDoc AND ItensPedCompraTodos.NumIntOrigem = CotacaoItemConcorrenciaTodos.NumIntDoc AND ItensPedCompraTodos.TipoOrigem = ? AND CotacaoItemConcorrenciaTodos.ItemConcorrencia = ItensConcorrenciaTodos.NumIntDoc AND ItensConcorrenciaTodos.Concorrencia = ConcorrenciaTodas.NumIntDoc AND ConcorrenciaTodas.NumIntDoc = ?", iMoeda, lCodigo, sProduto, dQuantidade, sUM, TIPO_ORIGEM_COTACAOITEMCONC, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67733
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67734
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
                            
        Set objItemPCInfo = New ClassItemPedCompraInfo
        
        objItemPCInfo.lPedCompra = lCodigo
        objItemPCInfo.sUM = sUM
        objItemPCInfo.dQuantReceber = dQuantidade
        objItemPCInfo.sProduto = sProduto
        objItemPCInfo.iMoeda = iMoeda
        
        'Adiciona novo item de Pedido de Compras
        colPedidos.Add objItemPCInfo
        
        'Busca próxima Requisição
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67735
    
    Loop
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    PedidoCompraTodos_Le_Concorrencia = SUCESSO
    
    Exit Function
    
Erro_PedidoCompraTodos_Le_Concorrencia:

    PedidoCompraTodos_Le_Concorrencia = gErr
    
    Select Case gErr
    
        Case 67732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 67733, 67734, 67735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRATODOS", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147638)
        
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TiposProduto_Le_Compravel(ByVal colCod_DescReduzida As AdmColCodigoNome) As Long
'Le todos os Codigos e DescReduzida dos tipos de produto que podem ser comprados (Compras=1) da tabela TiposDeProduto e coloca na colecao

Dim lComando As Long
Dim lErro As Long
Dim tTipoDeProduto As typeTipoDeProduto

On Error GoTo Erro_TiposProduto_Le_Compravel

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63117

    tTipoDeProduto.sDescricao = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)

    'Seleciona Tipos de Produto que podem ser comprados
    lErro = Comando_Executar(lComando, "SELECT TipoDeProduto, Descricao FROM TiposDeProduto WHERE  Compras= ? ORDER BY TipoDeProduto", tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao, 1)
    If lErro <> AD_SQL_SUCESSO Then Error 63118

    'Lê o primeiro Codigo e DescReduzida
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63119

    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o Codigo e a DescReduzida lidos na coleção
        colCod_DescReduzida.Add tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao

        'Lê o próximo Codigo e DescReduzida
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63120

    Loop

    'Fecha comando
    lErro = Comando_Fechar(lComando)

    TiposProduto_Le_Compravel = SUCESSO

    Exit Function

Erro_TiposProduto_Le_Compravel:

    TiposProduto_Le_Compravel = gErr

    Select Case gErr

        Case 63117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 63118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO1", gErr)

        Case 63119, 63120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", gErr, tTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147639)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Requisicoes_Le_GeracaoPC(ByVal gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Obtem as Requisicoes de Compras de acordo com a seleção determinada na tela

Dim lErro As Long
Dim sSelect As String
Dim lNumIntDoc As Long
Dim iFilialEmpresa As Integer
Dim lCodigo As Long
Dim dtData As Date
Dim dtDataLimite As Date
Dim lRequisitante As Long
Dim sCcl As String
Dim sNomeReduzido As String
Dim dtDataEnvio As Date
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim sOPCodigo As String
Dim lObservacao As Long
Dim iTipoDestino As Integer
Dim lFornCliDestino As Long
Dim lUrgente As Long
Dim iFilialDestino As Integer

On Error GoTo Erro_Requisicoes_Le_GeracaoPC

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 63956
    Next

    'Montar SELECT das Requisições dinamicamente
    Call Requisicoes_Le_GeracaoPC1(gobjGeracaoPedCompraReq, sSelect)

    sCcl = String(STRING_CCL, 0)
    sNomeReduzido = String(STRING_REQUISITANTE_NOMERED, 0)
    dtData = DATA_NULA
    dtDataLimite = DATA_NULA
    sOPCodigo = String(STRING_OPCODIGO, 0)

    'preparar parte fixa do SELECT
    lErro = Requisicoes_Le_GeracaoPC2(alComando(0), sSelect, sOPCodigo, dtDataEnvio, iFilialEmpresa, dtData, dtDataLimite, lCodigo, lRequisitante, sCcl, sNomeReduzido, iTipoDestino, lFornCliDestino, iFilialDestino, lNumIntDoc, lUrgente, lObservacao)
    If lErro <> SUCESSO Then gError 63957

    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das requisições
    lErro = Requisicoes_Le_GeracaoPC3(alComando(0), gobjGeracaoPedCompraReq)
    If lErro <> SUCESSO Then gError 63958

    'Processar os registros lidos
    lErro = Requisicoes_Le_GeracaoPC4(alComando(0), gobjGeracaoPedCompraReq, sOPCodigo, dtDataEnvio, iFilialEmpresa, dtData, dtDataLimite, lCodigo, lRequisitante, sCcl, sNomeReduzido, iTipoDestino, lFornCliDestino, iFilialDestino, lNumIntDoc, lUrgente, lObservacao)
    If lErro <> SUCESSO And lErro <> 68066 Then gError 63959

    'Se não encontrou nenhuma requisicao ==> erro
    If lErro = 68066 Then gError 63960

    'Fechar Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Requisicoes_Le_GeracaoPC = SUCESSO

    Exit Function

Erro_Requisicoes_Le_GeracaoPC:

    Requisicoes_Le_GeracaoPC = gErr

    Select Case gErr

        Case 63956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 63957, 63958, 63959
            'Erros tratados nas rotinas chamadas

        Case 63960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147640)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Sub Requisicoes_Le_GeracaoPC1(ByVal gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq, sSelect As String)
'Montar SELECT das Requisições dinamicamente

Dim sFields As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sGroupBy As String
Dim sHaving As String
Dim iIndice As Integer

    sSelect = "SELECT  MIN(RequisicaoCompra.OPCodigo)AS OPCodigo,  MIN(RequisicaoCompra.FilialEmpresa)AS FilialEmpresa, " _
              & "MIN(RequisicaoCompra.Data) AS Data, MIN(RequisicaoCompra.DataLimite)AS DataLimite," _
              & "MIN(RequisicaoCompra.Codigo)AS Codigo, MIN(RequisicaoCompra.Requisitante)AS Requisitante," _
              & "MIN(RequisicaoCompra.Ccl)AS Ccl, MIN(Requisitante.NomeReduzido)AS NomeReduzido,MIN(RequisicaoCompra.TipoDestino) AS TipoDestino, MIN (RequisicaoCompra.FornCliDestino) AS FornCliDestino, MIN(RequisicaoCompra.FilialDestino) AS FilialDestino, " _
              & "RequisicaoCompra.NumIntDoc, MIN(RequisicaoCompra.Urgente) AS Urgente, MIN(RequisicaoCompra.Observacao) AS Observacao"
    sFrom = " FROM RequisicaoCompra,Requisitante,Produtos, ItensReqCompra"
    
        sWhere = " WHERE ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND ItensReqCompra.Produto = Produtos.Codigo AND RequisicaoCompra.Requisitante = Requisitante.COdigo AND  RequisicaoCompra.DataEnvio <> ? AND Produtos.Compras = ?  AND (ItensReqCompra.Quantidade - ItensReqCompra.QuantCancelada - ItensReqCompra.QuantPedida - ItensReqCompra.QuantRecebida >= 0.0001)"
    
    If gobjGeracaoPedCompraReq.colTipoCategoria.Count > 0 Then
        sFrom = sFrom & ", ProdutoCategoria"
        sWhere = sWhere & " AND Produtos.Codigo = ProdutoCategoria.Produto "
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND RequisicaoCompra.FilialCompra = ?"
    If gobjGeracaoPedCompraReq.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data >=?"
    If gobjGeracaoPedCompraReq.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data <=?"
    If gobjGeracaoPedCompraReq.dtDataLimiteDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite >=?"
    If gobjGeracaoPedCompraReq.dtDataLimiteAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite <=?"
    If gobjGeracaoPedCompraReq.lCodigoDe <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo >= ? "
    If gobjGeracaoPedCompraReq.lCodigoAte <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo <= ? "
    
    If iIndice > 1 Then sWhere = sWhere & " )"
    
    If gobjGeracaoPedCompraReq.iSelecionaDestino <> 0 Then

        sWhere = sWhere & " AND RequisicaoCompra.TipoDestino = ?"

        If gobjGeracaoPedCompraReq.lFornCliDestino <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.FornCliDestino = ? "
        If gobjGeracaoPedCompraReq.iFilialDestino <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.FilialDestino = ? "

    End If
    
    For iIndice = 1 To gobjGeracaoPedCompraReq.colTipoProduto.Count

        If iIndice = 1 Then
            sWhere = sWhere & " AND Produtos.Tipo  IN ("
        Else
            sWhere = sWhere & ", "
        End If

        sWhere = sWhere & gobjGeracaoPedCompraReq.colTipoProduto(iIndice)

    Next

    If iIndice > 1 Then sWhere = sWhere & " )"
    
    If gobjGeracaoPedCompraReq.sCategoria <> "" Then sWhere = sWhere & " AND ProdutoCategoria.Categoria = '" & gobjGeracaoPedCompraReq.sCategoria & "'"

    For iIndice = 1 To gobjGeracaoPedCompraReq.colTipoCategoria.Count

        If iIndice = 1 Then
            sWhere = sWhere & " AND ProdutoCategoria.Item IN("
        Else
            sWhere = sWhere & ", "
        End If

        sWhere = sWhere & "  '" & gobjGeracaoPedCompraReq.colTipoCategoria(iIndice) & "'"

    Next
    
    If iIndice > 1 Then sWhere = sWhere & " )"
    
    sGroupBy = " GROUP BY RequisicaoCompra.NumIntDoc"
    sHaving = ""
    sSelect = sSelect & sFrom & sWhere & sGroupBy & sHaving

    Exit Sub

End Sub

Function Requisicoes_Le_GeracaoPC2(ByVal lComando As Long, ByVal sSelect As String, vsOPCodigo As Variant, vdtDataEnvio As Variant, viFilialEmpresa As Variant, vdtData As Variant, vdtDataLimite As Variant, vlCodigo As Variant, vlRequisitante As Variant, vsCcl As Variant, vsNomeReduzido As Variant, viTipoDestino As Variant, vlFornCliDestino As Variant, viFilialDestino As Variant, vlNumIntDoc As Variant, vlUrgente As Variant, vlObservacao As Variant) As Long

Dim ret As Integer, lErro As Long

On Error GoTo Erro_Requisicoes_Le_GeracaoPC2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 68067

    ret = Comando_BindVarInt(lComando, vsOPCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 178874

    ret = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then gError 68068

    ret = Comando_BindVarInt(lComando, vdtData)
    If (ret <> AD_SQL_SUCESSO) Then gError 68069

    ret = Comando_BindVarInt(lComando, vdtDataLimite)
    If (ret <> AD_SQL_SUCESSO) Then gError 68070

    ret = Comando_BindVarInt(lComando, vlCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 68071

    ret = Comando_BindVarInt(lComando, vlRequisitante)
    If (ret <> AD_SQL_SUCESSO) Then gError 68072

    ret = Comando_BindVarInt(lComando, vsCcl)
    If (ret <> AD_SQL_SUCESSO) Then gError 68073

    ret = Comando_BindVarInt(lComando, vsNomeReduzido)
    If (ret <> AD_SQL_SUCESSO) Then gError 68074
    
    ret = Comando_BindVarInt(lComando, viTipoDestino)
    If (ret <> AD_SQL_SUCESSO) Then gError 68075
    
    ret = Comando_BindVarInt(lComando, vlFornCliDestino)
    If (ret <> AD_SQL_SUCESSO) Then gError 68075

    ret = Comando_BindVarInt(lComando, viFilialDestino)
    If (ret <> AD_SQL_SUCESSO) Then gError 68075

    ret = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (ret <> AD_SQL_SUCESSO) Then gError 68075

    ret = Comando_BindVarInt(lComando, vlUrgente)
    If (ret <> AD_SQL_SUCESSO) Then gError 68071
    
    ret = Comando_BindVarInt(lComando, vlObservacao)
    If (ret <> AD_SQL_SUCESSO) Then gError 68075
    
    Requisicoes_Le_GeracaoPC2 = SUCESSO
    
    Exit Function

Erro_Requisicoes_Le_GeracaoPC2:

    Requisicoes_Le_GeracaoPC2 = gErr

    Select Case gErr

        Case 68067 To 68075, 178874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147641)

    End Select

    Exit Function

End Function

Function Requisicoes_Le_GeracaoPC3(ByVal lComando As Long, ByVal gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long

Dim lErro As Long
Dim lNumIntDoc As Variant
Dim viFilialEmpresa As Variant
Dim lCodigo As Variant
Dim dtDataEnvio As Variant
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant
Dim vdtDataLimiteDe As Variant
Dim vdtDataLimiteAte As Variant
Dim sCcl As Variant
Dim sNomeReduzido As Variant
Dim vdtDataEnvio As Variant
Dim viTipoDestino As Variant
Dim vlFornCliDestino As Variant
Dim viFilialDestino As Variant
Dim viCompras As Variant
Dim vsCategoria As String

On Error GoTo Erro_Requisicoes_Le_GeracaoPC3
            
    vdtDataEnvio = DATA_NULA
    lErro = Comando_BindVarInt(lComando, vdtDataEnvio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68076

    viCompras = PRODUTO_COMPRAVEL
    lErro = Comando_BindVarInt(lComando, viCompras)
    If (lErro <> AD_SQL_SUCESSO) Then gError 75397

    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68077
    End If

    If gobjGeracaoPedCompraReq.dtDataDe <> DATA_NULA Then
        vdtDataDe = gobjGeracaoPedCompraReq.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68078
    End If

    If gobjGeracaoPedCompraReq.dtDataAte <> DATA_NULA Then
        vdtDataAte = gobjGeracaoPedCompraReq.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68079
    End If

    If gobjGeracaoPedCompraReq.dtDataLimiteDe <> DATA_NULA Then
        vdtDataLimiteDe = gobjGeracaoPedCompraReq.dtDataLimiteDe
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68080
    End If

    If gobjGeracaoPedCompraReq.dtDataLimiteAte <> DATA_NULA Then
        vdtDataLimiteAte = gobjGeracaoPedCompraReq.dtDataLimiteAte
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68081
    End If

    If gobjGeracaoPedCompraReq.lCodigoDe <> 0 Then
        vlCodigoDe = gobjGeracaoPedCompraReq.lCodigoDe
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68082
    End If

    If gobjGeracaoPedCompraReq.lCodigoAte <> 0 Then
        vlCodigoAte = gobjGeracaoPedCompraReq.lCodigoAte
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 68083
    End If

    'Se foi seleciondo um Tipo de Destino
    If gobjGeracaoPedCompraReq.iSelecionaDestino <> 0 Then

        viTipoDestino = gobjGeracaoPedCompraReq.iTipoDestino
        lErro = Comando_BindVarInt(lComando, viTipoDestino)
        If (lErro <> AD_SQL_SUCESSO) Then gError 62686

        If gobjGeracaoPedCompraReq.lFornCliDestino <> 0 Then
            vlFornCliDestino = gobjGeracaoPedCompraReq.lFornCliDestino
            lErro = Comando_BindVarInt(lComando, vlFornCliDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 62687
        End If

        If gobjGeracaoPedCompraReq.iFilialDestino <> 0 Then
            viFilialDestino = gobjGeracaoPedCompraReq.iFilialDestino
            lErro = Comando_BindVarInt(lComando, viFilialDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 62690
        End If
        
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68084

    Requisicoes_Le_GeracaoPC3 = SUCESSO

    Exit Function

Erro_Requisicoes_Le_GeracaoPC3:

    Requisicoes_Le_GeracaoPC3 = gErr

    Select Case gErr

        Case 68076 To 68084, 62686, 62687, 62690, 75397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147642)

    End Select

    Exit Function

End Function

Function Requisicoes_Le_GeracaoPC4(ByVal lComando As Long, ByVal gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq, sOPCodigo As String, dtDataEnvio As Date, iFilialEmpresa As Integer, dtData As Date, dtDataLimite As Date, lCodigo As Long, lRequisitante As Long, sCcl As String, sNomeReduzido As String, iTipoDestino As Integer, lFornCliDestino As Long, iFilialDestino As Integer, lNumIntDoc As Long, lUrgente As Long, lObservacao As Long) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras
Dim iContador As Integer
Dim objCodNome As AdmlCodigoNome
Dim colObservacao As New AdmCollCodigoNome
Dim iPosicao As Integer
Dim objObs As New ClassObservacao

On Error GoTo Erro_Requisicoes_Le_GeracaoPC4

    Set gobjGeracaoPedCompraReq.colRequisicao = New Collection

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68077

    'Se não encontrou==>Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 68066

    iContador = 0

    Do While lErro = AD_SQL_SUCESSO

        iContador = iContador + 1

        'Verifica se o número de requisicoes é superior ao número máximo de requisicoes permitido
        If iContador = NUM_MAX_REQUISICOES Then gError 68078

        Set objReqCompras = New ClassRequisicaoCompras

        'Preenche objRequisicaoCompras com os dados lidos do BD
        objReqCompras.iSelecionado = MARCADO
        objReqCompras.dtData = dtData
        objReqCompras.dtDataLimite = dtDataLimite
        objReqCompras.iFilialEmpresa = iFilialEmpresa
        objReqCompras.lCodigo = lCodigo
        objReqCompras.lNumIntDoc = lNumIntDoc
        objReqCompras.lRequisitante = lRequisitante
        objReqCompras.dtDataEnvio = dtDataEnvio
        objReqCompras.sCcl = sCcl
        objReqCompras.iTipoDestino = iTipoDestino
        objReqCompras.iFilialDestino = iFilialDestino
        objReqCompras.lFornCliDestino = lFornCliDestino
        objReqCompras.lUrgente = lUrgente
        objReqCompras.lObservacao = lObservacao
        objReqCompras.sOPCodigo = sOPCodigo
        
        If lObservacao > 0 Then
            
            Call Busca_Na_Colecao(colObservacao, lObservacao, iPosicao)
            
            If iPosicao = 0 Then
                
                objObs.lNumInt = lObservacao
                
                lErro = CF("Observacao_Le", objObs)
                If lErro <> SUCESSO And lErro <> 53827 Then gError 62794
                If lErro <> SUCESSO Then Error 62795
                
                objReqCompras.sObservacao = objObs.sObservacao
                colObservacao.Add lObservacao, objObs.sObservacao
            
            Else
                
                objReqCompras.sObservacao = colObservacao(iPosicao).sNome
            
            End If
        
        End If
        
        'Adiciona objReqCompras na coleção de Requisicoes de Compra
        gobjGeracaoPedCompraReq.colRequisicao.Add objReqCompras

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68079

    Loop

    Requisicoes_Le_GeracaoPC4 = SUCESSO

    Exit Function

Erro_Requisicoes_Le_GeracaoPC4:

    Requisicoes_Le_GeracaoPC4 = gErr

    Select Case gErr

        Case 62794, 68066
            'Erro tratado na rotina chamadora

        Case 62795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, lObservacao)
        
        Case 68077, 68079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objReqCompras.lCodigo)

        Case 68078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147643)

    End Select

    Exit Function

End Function

Function PedidoCotacao_Le(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPedidoCotacao As typePedidoCotacao

On Error GoTo Erro_PedidoCotacao_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 53667

    With tPedidoCotacao

        .sContato = String(STRING_CONTATO, 0)

        'Seleciona os campos de PedidoCotacao que aparecem na tela onde
        'o Codigo e a FilialEmpresa forem iguais ao do objPedidoCotacao
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Fornecedor, Filial, Contato, DataEmissao, Data, DataValidade, TipoFrete,CondPagtoPrazo  FROM PedidoCotacao WHERE Codigo = ? AND FilialEmpresa=?", .lNumIntDoc, .lFornecedor, .iFilial, .sContato, .dtDataEmissao, .dtData, .dtDataValidade, .iTipoFrete, .iCondPagtoPrazo, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 53668

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53669
        If lErro = AD_SQL_SEM_DADOS Then Error 53670

        'Armazena as informações lidas do bd no objeto
        objPedidoCotacao.lNumIntDoc = .lNumIntDoc
        objPedidoCotacao.lFornecedor = .lFornecedor
        objPedidoCotacao.iFilial = .iFilial
        objPedidoCotacao.sContato = .sContato
        objPedidoCotacao.dtDataEmissao = .dtDataEmissao
        objPedidoCotacao.dtData = .dtData
        objPedidoCotacao.dtDataValidade = .dtDataValidade
        objPedidoCotacao.iTipoFrete = .iTipoFrete
        objPedidoCotacao.iCondPagtoPrazo = .iCondPagtoPrazo

    End With

    Call Comando_Fechar(lComando)

    PedidoCotacao_Le = SUCESSO

    Exit Function

Erro_PedidoCotacao_Le:

    PedidoCotacao_Le = Err

    Select Case Err

        Case 53667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53668, 53669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", Err, objPedidoCotacao.lCodigo)

        Case 53670 'Não existe na tabela de PedidoCotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147644)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCotacao_Le(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê os ítens do Pedido de Cotação

Dim iIndice As Integer
Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItensPedCotacao As typeItemPedCotacao
Dim colAlocacoes As ColAlocacoesItemNF
Dim tItensCotacao As typeItemCotacao
Dim tCotacaoProduto As typePedidoCotacao
Dim objItemPedidoCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao

On Error GoTo Erro_ItensPedCotacao_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 53637

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 53659

    tItensPedCotacao.sProduto = String(STRING_PRODUTO, 0)
    tItensPedCotacao.sUM = String(STRING_UM_SIGLA, 0)
    tItensPedCotacao.sObservacao = String(STRING_ITEMPEDCOTACAO_OBSERVACAO, 0)

    'Pesquisa os ítens de PedidoCotacao
    lErro = Comando_Executar(lComando, "SELECT ItemPedCotacao.Observacao, ItemPedCotacao.NumIntDoc, ItemPedCotacao.CotacaoProduto,  CotacaoProduto.NumIntDoc, CotacaoProduto.Produto, CotacaoProduto.Quantidade, CotacaoProduto.UM, CotacaoProduto.Fornecedor, CotacaoProduto.Filial, CotacaoProduto.Cotacao FROM PedidoCotacao, ItemPedCotacao, CotacaoProduto WHERE PedidoCotacao.NumIntDoc=ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND  PedidoCotacao.Codigo=? AND PedidoCotacao.FilialEmpresa=?", _
        tItensPedCotacao.sObservacao, tItensPedCotacao.lNumIntDoc, tItensPedCotacao.lCotacaoProduto, tCotacaoProduto.lNumIntDoc, tItensPedCotacao.sProduto, tItensPedCotacao.dQuantidade, tItensPedCotacao.sUM, tCotacaoProduto.lFornecedor, tCotacaoProduto.iFilial, tItensPedCotacao.lCotacaoProduto, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 53638
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53639

    Do While lErro <> AD_SQL_SEM_DADOS

        'Guarda na coleção de ítens de objPedidoCotacao
        Set objItemPedidoCotacao = New ClassItemPedCotacao

        objItemPedidoCotacao.lNumIntDoc = tItensPedCotacao.lNumIntDoc
        objPedidoCotacao.lCotacao = tItensPedCotacao.lCotacaoProduto
        objItemPedidoCotacao.sProduto = tItensPedCotacao.sProduto
        objItemPedidoCotacao.dQuantidade = tItensPedCotacao.dQuantidade
        objItemPedidoCotacao.sUM = tItensPedCotacao.sUM
        objItemPedidoCotacao.sObservacao = tItensPedCotacao.sObservacao

        lErro = Comando_Executar(lComando1, "SELECT CondPagto,PrecoUnitario, OutrasDespesas, ValorFrete, ValorSeguro, ValorDesconto, ValorTotal, ValorIPI, AliquotaIPI, AliquotaICMS, QuantEntrega, Observacao, PrazoEntrega, Moeda, Taxa FROM ItensCotacao WHERE ItemPedCotacao=?", tItensCotacao.iCondPagto, tItensCotacao.dPrecoUnitario, tItensCotacao.dOutrasDespesas, tItensCotacao.dValorFrete, tItensCotacao.dValorSeguro, tItensCotacao.dValorDesconto, tItensCotacao.dValorTotal, tItensCotacao.dValorIPI, tItensCotacao.dAliquotaIPI, tItensCotacao.dAliquotaICMS, tItensCotacao.dQuantEntrega, tItensCotacao.lObservacao, tItensCotacao.iPrazoEntrega, tItensCotacao.iMoeda, tItensCotacao.dTaxa, objItemPedidoCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 53650

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53651

        Do While lErro <> AD_SQL_SEM_DADOS

            'Guarda na coleção de itens cotação de objItemPedidoCotacao
            Set objItensCotacao = New ClassItemCotacao

            objItensCotacao.iCondPagto = tItensCotacao.iCondPagto
            objItensCotacao.dPrecoUnitario = tItensCotacao.dPrecoUnitario
            objItensCotacao.dOutrasDespesas = tItensCotacao.dOutrasDespesas
            objItensCotacao.dValorFrete = tItensCotacao.dValorFrete
            objItensCotacao.dValorSeguro = tItensCotacao.dValorSeguro
            objItensCotacao.dValorDesconto = tItensCotacao.dValorDesconto
            objItensCotacao.dValorTotal = tItensCotacao.dValorTotal
            objItensCotacao.dValorIPI = tItensCotacao.dValorIPI
            objItensCotacao.dAliquotaIPI = tItensCotacao.dAliquotaIPI
            objItensCotacao.dAliquotaICMS = tItensCotacao.dAliquotaICMS
            objItensCotacao.dQuantEntrega = tItensCotacao.dQuantEntrega
            objItensCotacao.lObservacao = tItensCotacao.lObservacao
            objItensCotacao.iPrazoEntrega = tItensCotacao.iPrazoEntrega
            objItensCotacao.iMoeda = tItensCotacao.iMoeda
            objItensCotacao.dTaxa = tItensCotacao.dTaxa

            objItemPedidoCotacao.colItensCotacao.Add objItensCotacao

            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53640

        Loop

        objPedidoCotacao.colItens.Add objItemPedidoCotacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53652

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPedCotacao_Le = SUCESSO

    Exit Function

Erro_ItensPedCotacao_Le:

    ItensPedCotacao_Le = Err

    Select Case Err

        Case 53637, 53659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53638, 53639, 53640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO", Err, objPedidoCotacao.lCodigo)

        Case 53650, 53651, 53652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147645)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Cotacao_Le(ByVal objCotacao As ClassCotacao) As Long

'Lê a tabela Cotacao
Dim lErro As Long
Dim lComando As Long
Dim tCotacao As typeCotacao

On Error GoTo Erro_Cotacao_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 53646

    With tCotacao

        tCotacao.sDescricao = String(STRING_PEDIDO_COTACAO_DESCRICAO, 0)

        'Seleciona os os campos da tabela Cotacao onde objCotacao.lNumIntDoc=objItemPedCotacao.lCotacao
        lErro = Comando_Executar(lComando, "SELECT  FilialEmpresa, Codigo, Descricao, Data, TipoDestino, FornCliDestino, FilialDestino, Comprador FROM Cotacao WHERE NumIntDoc = ? ", tCotacao.iFilialEmpresa, tCotacao.lCodigo, tCotacao.sDescricao, tCotacao.dtData, tCotacao.iTipoDestino, tCotacao.lFornCliDestino, tCotacao.iFilialDestino, tCotacao.iComprador, objCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 53647

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53648

        If lErro = AD_SQL_SEM_DADOS Then Error 53649

        'Armazena as informações lidas do bd no objeto
        objCotacao.iFilialEmpresa = tCotacao.iFilialEmpresa
        objCotacao.lCodigo = tCotacao.lCodigo
        objCotacao.sDescricao = tCotacao.sDescricao
        objCotacao.dtData = tCotacao.dtData
        objCotacao.iTipoDestino = tCotacao.iTipoDestino
        objCotacao.lFornCliDestino = tCotacao.lFornCliDestino
        objCotacao.iFilialDestino = tCotacao.iFilialDestino
        objCotacao.iComprador = tCotacao.iComprador

    End With

    Call Comando_Fechar(lComando)

    Cotacao_Le = SUCESSO

    Exit Function

Erro_Cotacao_Le:

    Cotacao_Le = Err

    Select Case Err

        Case 53646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53647, 53648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", Err)

        Case 53649 'Não existe na tabela Cotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147646)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cotacao_ObjetosAssociados_Lock(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Lê e locka os objetos associados a Cotacao

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iCodigo As Integer
Dim objCotacaoProduto As New ClassCotacaoProduto
Dim objProduto As New ClassProduto
Dim objUnidadeMedida As New ClassUnidadeDeMedida
Dim iIndice As Integer
Dim iCodFilial As Integer
Dim lCodFornecedor As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_Cotacao_ObjetosAssociados_Lock

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 63182
    Next

    If objCotacao.colCondPagtos.Count > 1 Then
        iCodigo = objCotacao.colCondPagtos.Item(2)

        'Lê e faz lock-shared na CondPagto a prazo
        lErro = CF("CondicaoPagto_Lock", alComando(1), iCodigo)
        If lErro <> SUCESSO And lErro <> 16361 Then Error 63209

    End If
    
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto

        objProduto.sCodigo = objCotacaoProduto.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 63273

        If lErro = 28030 Then Error 63275

        ' faz lock shared nos Produtos envolvidos
        lErro = CF("Produto_Lock_Shared", alComando(2), objProduto)
        If lErro <> SUCESSO And lErro <> 55235 Then Error 63210

        If lErro = 55235 Then Error 63278

    Next
    For Each objPedidoCotacao In colPedidoCotacao

        objFilialFornecedor.iCodFilial = objPedidoCotacao.iFilial
        objFilialFornecedor.lCodFornecedor = objPedidoCotacao.lFornecedor

        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 63279
        If lErro = 12929 Then Error 63280

        'Lê e faz lock shared nas FiliaisFornecedores envolvidas
        lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18900 Then Error 63238


    Next

    If objCotacao.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

        iCodFilial = objCotacao.iFilialDestino
        lCodFornecedor = objCotacao.lFornCliDestino

        'Lê e locka a Filial do Fornecedor
        lErro = CF("FilialFornecedor_Lock", lCodFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 42870 Then Error 63235

        If lErro = 42870 Then Error 63236

    End If

    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cotacao_ObjetosAssociados_Lock = SUCESSO

    Exit Function

Erro_Cotacao_ObjetosAssociados_Lock:

    Cotacao_ObjetosAssociados_Lock = Err

    Select Case Err

        Case 63182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63209, 63210, 63235, 63238, 63273, 63279

        Case 63234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, iCodFilial, lCodFornecedor)
        
        Case 63236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objCotacao.lFornCliDestino, objCotacao.iFilialDestino)
            
        Case 63275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 63278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", Err)

        Case 63280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147647)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ReqCompras_Le_GeracaoCotacao(ByVal objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Lê as Requisicoes de Compras e seus Itens a partir da selecao utilizada pelo usuário no tabSelecao

Dim lErro As Long
Dim tRequisicaoCompra As typeRequisicaoCompras
Dim lComando As Long
Dim lComando2 As Long
Dim sSelect As String
Dim sSelect2 As String
Dim tItemReqCompras As typeItemReqCompra
Dim tProduto As typeProduto
Dim sOPCodigo As String

On Error GoTo Erro_ReqCompras_Le_GeracaoCotacao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52925

    'Montar o SELECT dinamicamente
    Call ReqCompras_Le_GeracaoCotacao1(objGeracaoCotacao, sSelect)

    tRequisicaoCompra.sCcl = String(STRING_CCL, 0)
    
    With tItemReqCompras

        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        sOPCodigo = String(STRING_OPCODIGO, 0)
    
    End With

    'Preparar parte fixa do SELECT
    lErro = ReqCompras_Le_GeracaoCotacao2(lComando, sSelect, sOPCodigo, tRequisicaoCompra.lCodigo, tRequisicaoCompra.iFilialEmpresa, tRequisicaoCompra.dtData, tRequisicaoCompra.dtDataLimite, tRequisicaoCompra.iTipoDestino, tRequisicaoCompra.lUrgente, tRequisicaoCompra.lRequisitante, tRequisicaoCompra.sCcl, tRequisicaoCompra.lObservacao, tItemReqCompras.lNumIntDoc, tItemReqCompras.lReqCompra, tItemReqCompras.sProduto, tItemReqCompras.sDescProduto, tItemReqCompras.iStatus, tItemReqCompras.dQuantidade, tItemReqCompras.dQuantPedida, tItemReqCompras.dQuantRecebida, tItemReqCompras.dQuantCancelada, tItemReqCompras.sUM, tItemReqCompras.sCcl, tItemReqCompras.iAlmoxarifado, tItemReqCompras.sContaContabil, tItemReqCompras.iTipoTributacao, tItemReqCompras.lObservacao, tItemReqCompras.lFornecedor, tItemReqCompras.iFilial, tItemReqCompras.iExclusivo, tProduto.iTipo)
    If lErro <> SUCESSO Then Error 52926

    'Complementa a passagem dos parâmetros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das Requisições
    lErro = ReqCompras_Le_GeracaoCotacao3(lComando, objGeracaoCotacao)
    If lErro <> SUCESSO Then Error 52927

    'Processar os registros lidos
    lErro = ReqCompras_Le_GeracaoCotacao4(lComando, objGeracaoCotacao, sOPCodigo, tRequisicaoCompra.lCodigo, tRequisicaoCompra.iFilialEmpresa, tRequisicaoCompra.dtData, tRequisicaoCompra.dtDataLimite, tRequisicaoCompra.iTipoDestino, tRequisicaoCompra.lUrgente, tRequisicaoCompra.lRequisitante, tRequisicaoCompra.sCcl, tRequisicaoCompra.lObservacao, tItemReqCompras.lNumIntDoc, tItemReqCompras.lReqCompra, tItemReqCompras.sProduto, tItemReqCompras.sDescProduto, tItemReqCompras.iStatus, tItemReqCompras.dQuantidade, tItemReqCompras.dQuantPedida, tItemReqCompras.dQuantRecebida, tItemReqCompras.dQuantCancelada, tItemReqCompras.sUM, tItemReqCompras.sCcl, tItemReqCompras.iAlmoxarifado, tItemReqCompras.sContaContabil, tItemReqCompras.iTipoTributacao, tItemReqCompras.lObservacao, tItemReqCompras.lFornecedor, tItemReqCompras.iFilial, tItemReqCompras.iExclusivo, tProduto.iTipo)
    If lErro <> SUCESSO And lErro <> 52946 Then Error 52928

    'se não encontrou  - - - - - > Erro
    If lErro = 52946 Then Error 52929

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ReqCompras_Le_GeracaoCotacao = SUCESSO

    Exit Function

Erro_ReqCompras_Le_GeracaoCotacao:

    ReqCompras_Le_GeracaoCotacao = Err

    Select Case Err

        Case 52925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52926, 52927, 52928
            'Tratados Nas rotinas chmadas

        Case 52929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_ENCONTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147648)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub ReqCompras_Le_GeracaoCotacao1(ByVal objGeracaoCotacao As ClassGeracaoCotacao, sSelect As String)
'Monta Select Dinamicamente

Dim sFields As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sGroupBy As String

    sFields = " RequisicaoCompra.OPCodigo, RequisicaoCompra.Codigo, RequisicaoCompra.FilialEmpresa, RequisicaoCompra.Data, RequisicaoCompra.DataLimite ,RequisicaoCompra.TipoDestino ,RequisicaoCompra.Urgente , RequisicaoCompra.Requisitante , RequisicaoCompra.Ccl, RequisicaoCompra.Observacao"
    sFields = sFields & ", ItensReqCompra.NumIntDoc, ItensReqCompra.ReqCompra, ItensReqCompra.Produto, ItensReqCompra.DescProduto, ItensReqCompra.Status, ItensReqCompra.Quantidade, ItensReqCompra.QuantPedida, ItensReqCompra.QuantRecebida, ItensReqCompra.QuantCancelada, ItensReqCompra.UM, ItensReqCompra.Ccl, ItensReqCompra.Almoxarifado, ItensReqCompra.ContaContabil, ItensReqCompra.TipoTributacao, ItensReqCompra.Observacao, ItensReqCompra.Fornecedor, ItensReqCompra.Filial, ItensReqCompra.Exclusivo, Produtos.Tipo"
    sFrom = " FROM RequisicaoCompra, ItensReqCompra, Produtos"
    sWhere = " WHERE (ItensReqCompra.Quantidade-ItensReqCompra.QuantCancelada-ItensReqCompra.QuantPedida-ItensReqCompra.QuantRecebida)> 0 AND RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra AND ItensReqCompra.Produto = Produtos.Codigo AND RequisicaoCompra.FilialCompra = ? AND RequisicaoCompra.DataEnvio <> ? AND ItensReqCompra.Status = ?"
    sOrderBy = " ORDER BY " & objGeracaoCotacao.sOrdenacaoReq & ",ItensReqCompra.NumIntDoc"

    If objGeracaoCotacao.lCodigoDe <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo >= ?"

    If objGeracaoCotacao.lCodigoAte <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.Codigo <= ?"

    If objGeracaoCotacao.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data >= ?"

    If objGeracaoCotacao.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.Data <= ?"

    If objGeracaoCotacao.dtDataLimiteDe <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite >= ?"

    If objGeracaoCotacao.dtDataLimiteAte <> DATA_NULA Then sWhere = sWhere & " AND RequisicaoCompra.DataLimite <= ?"

    If objGeracaoCotacao.iTipoDestino <> TIPO_DESTINO_AUSENTE Then

        sWhere = sWhere & " AND RequisicaoCompra.TipoDestino = ?"

        If objGeracaoCotacao.lFornCliDestino <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.FornCliDestino = ? "
        If objGeracaoCotacao.iFilialDestino <> 0 Then sWhere = sWhere & " AND RequisicaoCompra.FilialDestino = ? "

    End If
    
    If objGeracaoCotacao.iExibeReqCotadas = NAO_EXIBE_REQUISICOES_COTADAS Then
        sWhere = sWhere & " AND ItensReqCompra.NumIntDoc NOT IN (SELECT CotacaoProdutoItemRC.ItemReqCompra FROM CotacaoProdutoItemRC)"
    End If
    
    sSelect = "SELECT" & sFields & sFrom & sWhere & sOrderBy

End Sub

Function ReqCompras_Le_GeracaoCotacao2(ByVal lComando As Long, ByVal sSelect As String, vOPCodigo As Variant, vCodigo As Variant, vFilialEmpresa As Variant, vData As Variant, vDataLimite As Variant, vTipoDestino As Variant, vUrgente As Variant, vRequisitante As Variant, vCcl As Variant, vObservacao As Variant, vNumIntDoc As Variant, vReqCompra As Variant, vProduto As Variant, vDescProduto As Variant, vStatus As Variant, vQuantidade As Variant, vQuantPedida As Variant, vQuantRecebida As Variant, vQuantCancelada As Variant, vUM As Variant, vCclItem As Variant, vAlmoxarifado As Variant, vContaContabil As Variant, vTipoTributacao As Variant, vObservacaoItem As Variant, vFornecedor As Variant, vFilial As Variant, vExclusivo As Variant, vTipoProduto As Variant) As Long

Dim lErro As Long
Dim ret As Integer

On Error GoTo Erro_ReqCompras_Le_GeracaoCotacao2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 52929

    ret = Comando_BindVarInt(lComando, vOPCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 178873

    ret = Comando_BindVarInt(lComando, vCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 52930

    ret = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then gError 64316

    ret = Comando_BindVarInt(lComando, vData)
    If (ret <> AD_SQL_SUCESSO) Then gError 52931

    ret = Comando_BindVarInt(lComando, vDataLimite)
    If (ret <> AD_SQL_SUCESSO) Then gError 52932

    ret = Comando_BindVarInt(lComando, vTipoDestino)
    If (ret <> AD_SQL_SUCESSO) Then gError 52948

    ret = Comando_BindVarInt(lComando, vUrgente)
    If (ret <> AD_SQL_SUCESSO) Then gError 52933

    ret = Comando_BindVarInt(lComando, vRequisitante)
    If (ret <> AD_SQL_SUCESSO) Then gError 52934

    ret = Comando_BindVarInt(lComando, vCcl)
    If (ret <> AD_SQL_SUCESSO) Then gError 52935

    ret = Comando_BindVarInt(lComando, vObservacao)
    If (ret <> AD_SQL_SUCESSO) Then gError 52936

    ret = Comando_BindVarInt(lComando, vNumIntDoc)
    If (ret <> AD_SQL_SUCESSO) Then gError 25979
    
    ret = Comando_BindVarInt(lComando, vReqCompra)
    If (ret <> AD_SQL_SUCESSO) Then gError 25980
 
    ret = Comando_BindVarInt(lComando, vProduto)
    If (ret <> AD_SQL_SUCESSO) Then gError 25981

    ret = Comando_BindVarInt(lComando, vDescProduto)
    If (ret <> AD_SQL_SUCESSO) Then gError 25982

    ret = Comando_BindVarInt(lComando, vStatus)
    If (ret <> AD_SQL_SUCESSO) Then gError 25983

    ret = Comando_BindVarInt(lComando, vQuantidade)
    If (ret <> AD_SQL_SUCESSO) Then gError 25984

    ret = Comando_BindVarInt(lComando, vQuantPedida)
    If (ret <> AD_SQL_SUCESSO) Then gError 25985

    ret = Comando_BindVarInt(lComando, vQuantRecebida)
    If (ret <> AD_SQL_SUCESSO) Then gError 25986
    
    ret = Comando_BindVarInt(lComando, vQuantCancelada)
    If (ret <> AD_SQL_SUCESSO) Then gError 25987
    
    ret = Comando_BindVarInt(lComando, vUM)
    If (ret <> AD_SQL_SUCESSO) Then gError 25988
    
    ret = Comando_BindVarInt(lComando, vCclItem)
    If (ret <> AD_SQL_SUCESSO) Then gError 25989
    
    ret = Comando_BindVarInt(lComando, vAlmoxarifado)
    If (ret <> AD_SQL_SUCESSO) Then gError 25990
    
    ret = Comando_BindVarInt(lComando, vContaContabil)
    If (ret <> AD_SQL_SUCESSO) Then gError 25991
    
    ret = Comando_BindVarInt(lComando, vTipoTributacao)
    If (ret <> AD_SQL_SUCESSO) Then gError 25992
    
    ret = Comando_BindVarInt(lComando, vObservacaoItem)
    If (ret <> AD_SQL_SUCESSO) Then gError 25993
    
    ret = Comando_BindVarInt(lComando, vFornecedor)
    If (ret <> AD_SQL_SUCESSO) Then gError 25994
    
    ret = Comando_BindVarInt(lComando, vFilial)
    If (ret <> AD_SQL_SUCESSO) Then gError 25995
    
    ret = Comando_BindVarInt(lComando, vExclusivo)
    If (ret <> AD_SQL_SUCESSO) Then gError 25996
    
    ret = Comando_BindVarInt(lComando, vTipoProduto)
    If (ret <> AD_SQL_SUCESSO) Then gError 25997
    
    ReqCompras_Le_GeracaoCotacao2 = SUCESSO

    Exit Function

Erro_ReqCompras_Le_GeracaoCotacao2:

    ReqCompras_Le_GeracaoCotacao2 = gErr

    Select Case gErr

        Case 52929 To 52936, 52948, 64316, 178873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)

        Case 25979 To 25996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
        
        Case 25997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147649)

    End Select

    Exit Function

End Function

Function ReqCompras_Le_GeracaoCotacao3(ByVal lComando As Long, ByVal objGeracaoCotacao As ClassGeracaoCotacao) As Long

Dim lErro As Long
Dim vdtDataAte As Variant
Dim vdtDataLimiteDe As Variant
Dim vdtDataLimiteAte As Variant
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim viTipoDestino As Variant
Dim vlFornCliDestino As Variant
Dim viFilialDestino As Variant
Dim vdtData As Variant
Dim viAberta As Variant
Dim vdtDataDe As Variant
Dim viCotadas As Variant
Dim viFilialCompra As Variant

On Error GoTo Erro_ReqCompras_Le_GeracaoCotacao3

    viFilialCompra = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, viFilialCompra)
    If lErro <> AD_SQL_SUCESSO Then Error 52956

    vdtData = DATA_NULA
    lErro = Comando_BindVarInt(lComando, vdtData)
    If lErro <> AD_SQL_SUCESSO Then Error 63569

    viAberta = ITEM_REQ_ABERTO
    lErro = Comando_BindVarInt(lComando, viAberta)
    If lErro <> AD_SQL_SUCESSO Then Error 52957

    If objGeracaoCotacao.lCodigoDe <> 0 Then
        vlCodigoDe = objGeracaoCotacao.lCodigoDe
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If lErro <> AD_SQL_SUCESSO Then Error 52940
    End If

    If objGeracaoCotacao.lCodigoAte <> 0 Then
        vlCodigoAte = objGeracaoCotacao.lCodigoAte
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If lErro <> AD_SQL_SUCESSO Then Error 52941
    End If

    If objGeracaoCotacao.dtDataDe <> DATA_NULA Then
        vdtDataDe = objGeracaoCotacao.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If lErro <> AD_SQL_SUCESSO Then Error 52938
    End If

    If objGeracaoCotacao.dtDataAte <> DATA_NULA Then
        vdtDataAte = objGeracaoCotacao.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If lErro <> AD_SQL_SUCESSO Then Error 52939
    End If

    If objGeracaoCotacao.dtDataLimiteDe <> DATA_NULA Then
        vdtDataDe = objGeracaoCotacao.dtDataLimiteDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If lErro <> AD_SQL_SUCESSO Then Error 63570
    End If

    If objGeracaoCotacao.dtDataLimiteAte <> DATA_NULA Then
        vdtDataAte = objGeracaoCotacao.dtDataLimiteAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If lErro <> AD_SQL_SUCESSO Then Error 63571
    End If

    If objGeracaoCotacao.iTipoDestino <> TIPO_DESTINO_AUSENTE Then
        viTipoDestino = objGeracaoCotacao.iTipoDestino
        lErro = Comando_BindVarInt(lComando, viTipoDestino)
        If lErro <> AD_SQL_SUCESSO Then Error 52942

        If objGeracaoCotacao.lFornCliDestino <> 0 Then
            vlFornCliDestino = objGeracaoCotacao.lFornCliDestino
            lErro = Comando_BindVarInt(lComando, vlFornCliDestino)
            If (lErro <> AD_SQL_SUCESSO) Then Error 52943
        End If

        If objGeracaoCotacao.iFilialDestino <> 0 Then
            viFilialDestino = objGeracaoCotacao.iFilialDestino
            lErro = Comando_BindVarInt(lComando, viFilialDestino)
            If (lErro <> AD_SQL_SUCESSO) Then Error 52944
        End If

    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 63607

    ReqCompras_Le_GeracaoCotacao3 = SUCESSO

    Exit Function

Erro_ReqCompras_Le_GeracaoCotacao3:

    ReqCompras_Le_GeracaoCotacao3 = Err

    Select Case Err

        Case 52937 To 52944, 52956, 52957, 63569, 63570, 63571, 63607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147650)

    End Select

    Exit Function

End Function

Function ReqCompras_Le_GeracaoCotacao4(ByVal lComando As Long, ByVal objGeracaoCotacao As ClassGeracaoCotacao, sOPCodigo As String, lCodigo As Long, iFilialEmpresa As Integer, dtData As Date, dtDataLimite As Date, iTipoDestino As Integer, lUrgente As Long, lRequisitante As Long, sCcl As String, lObservacao As Long, lNumIntDoc As Long, lReqCompra As Long, sProduto As String, sDescProduto As String, iStatus As Integer, dQuantidade As Double, dQuantPedida As Double, dQuantRecebida As Double, dQuantCancelada As Double, sUM As String, sCclItem As String, iAlmoxarifado As Integer, sContaContabil As String, iTipoTributacao As Integer, lObservacaoItem As Long, lFornecedor As Long, iFilial As Integer, iExclusivo As Integer, iTipoProduto As Integer) As Long
'Armazena em objGeracaoCotacao Requisicoes lidas e seus Itens

Dim lErro As Long
Dim objRequisicaoCompras As ClassRequisicaoCompras
Dim objItemReqCompras As ClassItemReqCompras
Dim lNumIntReqCompras As Long
Dim bTipoSelecionado As Boolean
Dim vTipoProduto As Variant
Dim iIndice As Integer

On Error GoTo Erro_ReqCompras_Le_GeracaoCotacao4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52945

    'Se não encontrou ReqCompras ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 52946

    Set objGeracaoCotacao.colReqCompra = New Collection
    
    'Armazena zero no num int da ReqCompras atual
    lNumIntReqCompras = 0

    Do While lErro = AD_SQL_SUCESSO

        If lReqCompra <> lNumIntReqCompras Then
                
            'Armazena o novo num int de ReqCompras
            lNumIntReqCompras = lReqCompra
            
            Set objRequisicaoCompras = New ClassRequisicaoCompras
    
            'Preenche objRequisicaoCompras com os dados lidos do BD
            objRequisicaoCompras.lCodigo = lCodigo
            objRequisicaoCompras.iFilialEmpresa = iFilialEmpresa
            objRequisicaoCompras.dtData = dtData
            objRequisicaoCompras.dtDataLimite = dtDataLimite
            objRequisicaoCompras.iTipoDestino = iTipoDestino
            objRequisicaoCompras.lUrgente = lUrgente
            objRequisicaoCompras.lRequisitante = lRequisitante
            objRequisicaoCompras.sCcl = sCcl
            objRequisicaoCompras.lObservacao = lObservacao
            objRequisicaoCompras.iSelecionado = MARCADO
            objRequisicaoCompras.sOPCodigo = sOPCodigo

            objGeracaoCotacao.colReqCompra.Add objRequisicaoCompras
            
        End If

        If objGeracaoCotacao.colTipoProduto.Count = 0 Then
            bTipoSelecionado = True
        Else
            bTipoSelecionado = False
        End If
            
        'Verifica se o Tipo de Produto do ItemRequisicao é um dos selecinados
        For Each vTipoProduto In objGeracaoCotacao.colTipoProduto
            
            If vTipoProduto = iTipoProduto Then
                bTipoSelecionado = True
            End If
        
        Next

        If bTipoSelecionado = True Then
        
            Set objItemReqCompras = New ClassItemReqCompras
            
            objItemReqCompras.lNumIntDoc = lNumIntDoc
            objItemReqCompras.iItem = objRequisicaoCompras.colItens.Count + 1
            objItemReqCompras.lReqCompra = lReqCompra
            objItemReqCompras.sProduto = sProduto
            objItemReqCompras.sDescProduto = sDescProduto
            objItemReqCompras.iStatus = iStatus
            objItemReqCompras.dQuantidade = dQuantidade
            objItemReqCompras.dQuantPedida = dQuantPedida
            objItemReqCompras.dQuantRecebida = dQuantRecebida
            objItemReqCompras.dQuantCancelada = dQuantCancelada
            objItemReqCompras.sUM = sUM
            objItemReqCompras.sCcl = sCclItem
            objItemReqCompras.iAlmoxarifado = iAlmoxarifado
            objItemReqCompras.sContaContabil = sContaContabil
            objItemReqCompras.iTipoTributacao = iTipoTributacao
            objItemReqCompras.lObservacao = lObservacaoItem
            objItemReqCompras.lFornecedor = lFornecedor
            objItemReqCompras.iFilial = iFilial
            objItemReqCompras.iExclusivo = iExclusivo
            objItemReqCompras.iSelecionado = 1
  
            objRequisicaoCompras.colItens.Add objItemReqCompras
            
        End If
                
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52947

    Loop

    For iIndice = objGeracaoCotacao.colReqCompra.Count To 1 Step -1
        If objGeracaoCotacao.colReqCompra(iIndice).colItens.Count = 0 Then
            objGeracaoCotacao.colReqCompra.Remove iIndice
        End If
    Next
    
    If objGeracaoCotacao.colReqCompra.Count = 0 Then Error 52946

    ReqCompras_Le_GeracaoCotacao4 = SUCESSO

    Exit Function

Erro_ReqCompras_Le_GeracaoCotacao4:

    ReqCompras_Le_GeracaoCotacao4 = Err

    Select Case Err

        Case 52945, 52947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", Err)

        Case 52946
            'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147651)

    End Select

    Exit Function

End Function

Function ItensReqCompras_Le(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras passada por objRequisiçãoCompras

Dim lErro As Long
Dim tItemReqCompras As typeItemReqCompra
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long
Dim iItem As Integer

On Error GoTo Erro_ItensReqCompras_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54398

    With tItemReqCompras

        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)

        'Lê o ítem da Requisição na tabela ItensReqCompras
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompra WHERE ReqCompra = ? ORDER BY NumIntDoc", _
            .lNumIntDoc, .lReqCompra, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iTipoTributacao, .lObservacao, .lFornecedor, .iFilial, .iExclusivo, objRequisicaoCompras.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 54399
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54400

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 54377

        iItem = 0
        
        'Equanto encontrar
        Do While lErro = AD_SQL_SUCESSO

            Set objItemReqCompras = New ClassItemReqCompras
            iItem = iItem + 1

            'Preenche o objItemReqCompras
            objItemReqCompras.lNumIntDoc = .lNumIntDoc
            objItemReqCompras.lReqCompra = .lReqCompra
            objItemReqCompras.sProduto = .sProduto
            objItemReqCompras.sDescProduto = .sDescProduto
            objItemReqCompras.iStatus = .iStatus
            objItemReqCompras.dQuantidade = .dQuantidade
            objItemReqCompras.dQuantPedida = .dQuantPedida
            objItemReqCompras.dQuantRecebida = .dQuantRecebida
            objItemReqCompras.dQuantCancelada = .dQuantCancelada
            objItemReqCompras.sUM = .sUM
            objItemReqCompras.sCcl = .sCcl
            objItemReqCompras.iAlmoxarifado = .iAlmoxarifado
            objItemReqCompras.sContaContabil = .sContaContabil
            objItemReqCompras.iTipoTributacao = .iTipoTributacao
            objItemReqCompras.lObservacao = .lObservacao
            objItemReqCompras.lFornecedor = .lFornecedor
            objItemReqCompras.iFilial = .iFilial
            objItemReqCompras.iExclusivo = .iExclusivo
            objItemReqCompras.iSelecionado = 1
            objItemReqCompras.iItem = iItem

            'Adiciona o objItemReqCompras na coleção colItens do objRequisiçãoCompras
            objRequisicaoCompras.colItens.Add objItemReqCompras

            'Busca no BD o próximo Ítem de Requisição de compras
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54401

        Loop

    End With

    Call Comando_Fechar(lComando)

    ItensReqCompras_Le = SUCESSO

    Exit Function

Erro_ItensReqCompras_Le:

    ItensReqCompras_Le = Err

    Select Case Err

        Case 54377

        Case 54398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54399, 54400, 54401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147652)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensReqCompras_Le_NaoCotadas(ByVal objReqCompra As ClassRequisicaoCompras) As Long
'Lê os Itens da Requisicao de Compras que não foram cotados e adicionar na
'coleção de Itens de objReqCompra

Dim lErro As Long
Dim tItemReqCompras As typeItemReqCompra
Dim objItemReqCompra As ClassItemReqCompras
Dim lComando As Long
Dim t As typeCotacaoProduto

On Error GoTo Erro_ItensReqCompras_Le_NaoCotadas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51288

    With tItemReqCompras

        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)

        'Le o item da Requisicao de Compra
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Produto,DescProduto,Status,Quantidade,QuantPedida,QuantRecebida,QuantCancelada,UM,Ccl,Almoxarifado,ContaContabil,CreditaICMS,CreditaIPI,Observacao,Fornecedor,Filial,Exclusivo FROM ItensReqCompra WHERE ReqCompra = ? AND ItensReqCompra.NumIntDoc NOT IN (SELECT CotacaoProdutoItemRC.ItemReqCompra FROM CotacaoProdutoItemRC) ORDER BY NumIntDoc", .lNumIntDoc, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iCreditaICMS, .iCreditaIPI, .lObservacao, .lFornecedor, .iFilial, .iExclusivo, objReqCompra.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 51289

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51290

        'If lErro = AD_SQL_SEM_DADOS Then Error 63609

        Do While lErro = AD_SQL_SUCESSO

            Set objItemReqCompra = New ClassItemReqCompras

            'Preenche objItemReqCompra com os dados lidos do BD
            objItemReqCompra.lNumIntDoc = .lNumIntDoc
            objItemReqCompra.sProduto = .sProduto
            objItemReqCompra.sDescProduto = .sDescProduto
            objItemReqCompra.iStatus = .iStatus
            objItemReqCompra.dQuantidade = .dQuantidade
            objItemReqCompra.dQuantPedida = .dQuantPedida
            objItemReqCompra.dQuantRecebida = .dQuantRecebida
            objItemReqCompra.dQuantCancelada = .dQuantCancelada
            objItemReqCompra.sUM = .sUM
            objItemReqCompra.sCcl = .sCcl
            objItemReqCompra.iAlmoxarifado = .iAlmoxarifado
            objItemReqCompra.sContaContabil = .sContaContabil
            objItemReqCompra.iCreditaICMS = .iCreditaICMS
            objItemReqCompra.iCreditaIPI = .iCreditaIPI
            objItemReqCompra.sObservacao = .sObservacao
            objItemReqCompra.lFornecedor = .lFornecedor
            objItemReqCompra.iFilial = .iFilial
            objItemReqCompra.iExclusivo = .iExclusivo

            'Adiciona o objItemReqCompra em colItens
            objReqCompra.colItens.Add objItemReqCompra

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51291

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensReqCompras_Le_NaoCotadas = SUCESSO

    Exit Function

Erro_ItensReqCompras_Le_NaoCotadas:

    ItensReqCompras_Le_NaoCotadas = Err

    Select Case Err

        Case 51288
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 51289, 51290, 51291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err)

        Case 63609
            'Sem dados ==>Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147653)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoProdutoFornecedor_Le(ByVal colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection) As Long
'Verifica se o fornecedor do Item é exclusivo ou não e devolve seus Fornecedores em colFornecedorProdutoFF

Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF

On Error GoTo Erro_CotacaoProdutoFornecedor_Le

    Set colFornecedorProdutoFF = New Collection

    For Each objCotacaoProduto In colCotacaoProduto

        'Se o Fornecedor é exclusivo
        If objCotacaoProduto.lFornecedor > 0 Then
            
            Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF
            
            objFornecedorProdutoFF.iEscolhido = Selecionado
            objFornecedorProdutoFF.iSelecionado = objCotacaoProduto.iSelecionado
            objFornecedorProdutoFF.lFornecedor = objCotacaoProduto.lFornecedor
            objFornecedorProdutoFF.iFilialForn = objCotacaoProduto.iFilial
            objFornecedorProdutoFF.sProduto = objCotacaoProduto.sProduto
            objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa

            'Lê a tabela de FornecedorProdutoFF a partir do Fornecedor,Filial,Produto e FilialEmpresa
            'passados como parametros
            lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 54217 Then gError 63587

            'Se não encontrou ==> erro
            If lErro = 54217 Then gError 63588
            
            'Referencia CotacaoProduto em objFornecedorProdutoFF
            objFornecedorProdutoFF.lNumIntCotacaoProduto = objCotacaoProduto.lNumIntDoc
            
            objFornecedorProdutoFF.lNumIntDoc = colFornecedorProdutoFF.Count + 1
            
            colFornecedorProdutoFF.Add objFornecedorProdutoFF
            
        'Se não tem Fornecedor exclusivo
        ElseIf objCotacaoProduto.lFornecedor = 0 Then

            'Lê os fornecedores associados ao produto passado como parametro,gerando itens
            'e armazenando estes itens em colCotacaoProdutoFornecedor
            lErro = CF("FornecedorProdutoFF_Le_CotacaoProdutoForn", objCotacaoProduto, colFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 63593 Then gError 63597

        End If

    Next

    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
    
        'Completa os campos q faltam do objFornecedor ProdutoFF
        'Lê as informações da última cotação
        lErro = CF("UltimaCotacao_Le_FornecedorProduto", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 86282
        
        lErro = CF("UltimaCompra_Le_FornecedorProduto", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 86283
            
    Next
    

    CotacaoProdutoFornecedor_Le = SUCESSO

    Exit Function

Erro_CotacaoProdutoFornecedor_Le:

    CotacaoProdutoFornecedor_Le = gErr

    Select Case gErr

        Case 63587, 63597, 86282, 86283
            'Erro tratado na rotina chamada

        Case 63588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.lFornecedor)

        Case 63598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.lFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147654)

    End Select

    Exit Function

End Function

'?? Tirar Private e voltar para ClassMATSelect que está com check-out
Function GeracaoReqPedVenda_Le_Pedidos(ByVal gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda) As Long
'Busca no BD os Pedidos de Venda que atendam as condicoes informadas no TabSelecao da tela

Dim lErro As Long, sSelect As String
Dim lComando As Long
Dim dtDataEntrega As Date
Dim iFilialEmpresa As Integer
Dim lCodigo As Long
Dim lCliente As Long
Dim iFilial As Integer
Dim dtEmissaoPedido As Date
Dim dValorTotal As Double
Dim dtDataNula As Date
Dim iSelect As Integer
Dim sNomeCliente As String
Dim iFilialEmpresa2 As Integer
Dim iFilialEmpresa3 As Integer


On Error GoTo Erro_GeracaoReqPedVenda_Le_Pedidos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 68154

    'Montar SELECT dos pedidos de venda dinamicamente
    Call GeracaoReqPedVenda_Le_Pedidos1(gobjGeracaoReqPedVenda, sSelect)

    sNomeCliente = String(STRING_BUFFER_MAX_TEXTO, 0)
    dtDataNula = DATA_NULA

    'Informa que os dados a serem preenchidos são do primeiro ou segundo
    'Preparar parte fixa do SELECT
    lErro = GeracaoReqPedVenda_Le_Pedidos2(lComando, sSelect, dtDataEntrega, sNomeCliente, iFilialEmpresa, lCodigo, lCliente, iFilial, dtEmissaoPedido, dValorTotal, dtDataNula, iFilialEmpresa2, iFilialEmpresa3, dtDataNula, iSelect)
    If lErro <> SUCESSO Then gError 68155

    'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao dos bloqueios
    lErro = GeracaoReqPedVenda_Le_Pedidos3(lComando, gobjGeracaoReqPedVenda, iSelect)
    If lErro <> SUCESSO Then gError 68156

    'Processar os registros lidos
    lErro = GeracaoReqPedVenda_Le_Pedidos4(lComando, gobjGeracaoReqPedVenda, dtDataEntrega, sNomeCliente, iFilialEmpresa, lCodigo, lCliente, iFilial, dtEmissaoPedido, dValorTotal)
    If lErro <> SUCESSO And lErro <> 68188 Then gError 68157

    'Se não encontrou nenhum Pedido ---> Erro
    If lErro = 68188 Then gError 68158

    Call Comando_Fechar(lComando)

    GeracaoReqPedVenda_Le_Pedidos = SUCESSO

    Exit Function

Erro_GeracaoReqPedVenda_Le_Pedidos:

    GeracaoReqPedVenda_Le_Pedidos = Err

    Select Case gErr

        Case 68155 To 68157

        Case 68154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_ENCONTRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147655)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub GeracaoReqPedVenda_Le_Pedidos1(ByVal gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String, sGroupBy As String, sHaving As String
Dim sUnion As String, sWhere2 As String, sHaving2 As String
Dim sWhereComum As String 'guarda a parte comum no WHERE dos 2 selects

    sFields = "DataEntrega, NomeCliente, FilialEmpresa, Codigo, Cliente, Filial, EmissaoPedido, ValorTotal "
    sFrom = " FROM PedidosVendaGeracaoReq"
    sWhere = " WHERE FilialEmpresa=? "
    sOrderBy = " ORDER BY " & gobjGeracaoReqPedVenda.sOrdenacao
    
    If gobjGeracaoReqPedVenda.iExibeTodos = vbUnchecked Then
        
        If gobjGeracaoReqPedVenda.PedidoDe <> 0 Then sWhereComum = sWhereComum & " AND Codigo >= ?"

        If gobjGeracaoReqPedVenda.PedidoAte <> 0 Then sWhereComum = sWhereComum & " AND Codigo <= ?"

        If gobjGeracaoReqPedVenda.ClienteDe <> 0 Then sWhereComum = sWhereComum & " AND Cliente >= ?"

        If gobjGeracaoReqPedVenda.ClienteAte <> 0 Then sWhereComum = sWhereComum & " AND Cliente <= ?"

        If gobjGeracaoReqPedVenda.DataEmissaoDe <> DATA_NULA Then sWhereComum = sWhereComum & " AND EmissaoPedido >= ?"

        If gobjGeracaoReqPedVenda.DataEmissaoAte <> DATA_NULA Then sWhereComum = sWhereComum & " AND EmissaoPedido <= ?"

        If gobjGeracaoReqPedVenda.DataEntregaDe <> DATA_NULA Then sWhereComum = sWhereComum & " AND DataEntrega >= ?"
        If gobjGeracaoReqPedVenda.DataEntregaAte <> DATA_NULA Then sWhereComum = sWhereComum & " AND DataEntrega <= ?"

    End If
    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sWhereComum & sGroupBy & sHaving

End Sub

'??? iSelect parece parametro inutil. jones 2002
Function GeracaoReqPedVenda_Le_Pedidos2(ByVal lComando As Long, ByVal sSelect As String, vDataEntrega As Variant, vNomeCliente As Variant, vFilialEmpresa As Variant, vCodigo As Variant, vCliente As Variant, vFilial As Variant, vEmissaoPedido As Variant, vValorTotal As Variant, vDataNula As Variant, vFilialEmpresa2 As Variant, vFilialEmpresa3 As Variant, vDataNula2 As Variant, iSelect As Integer) As Long
'Preparar parte fixa do SELECT

Dim lErro As Long

On Error GoTo Erro_GeracaoReqPedVenda_Le_Pedidos2

    'Inicializa as variaves de leitura
    lErro = Comando_PrepararInt(lComando, sSelect)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68165

    lErro = Comando_BindVarInt(lComando, vDataEntrega)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68166

    lErro = Comando_BindVarInt(lComando, vNomeCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68167

    lErro = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68168

    lErro = Comando_BindVarInt(lComando, vCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68169

    lErro = Comando_BindVarInt(lComando, vCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68170

    lErro = Comando_BindVarInt(lComando, vFilial)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68171

    lErro = Comando_BindVarInt(lComando, vEmissaoPedido)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68172

    lErro = Comando_BindVarInt(lComando, vValorTotal)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68173
        
    'Valores utilizados nas clausulas WHERE dos 2 SELECT's
    vFilialEmpresa2 = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, vFilialEmpresa2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68174
   
    GeracaoReqPedVenda_Le_Pedidos2 = SUCESSO

    Exit Function

Erro_GeracaoReqPedVenda_Le_Pedidos2:

    GeracaoReqPedVenda_Le_Pedidos2 = gErr

    Select Case gErr

        Case 68165 To 68175, 68218, 68219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_RC", gErr)


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147656)

    End Select

    Exit Function

End Function

'??? iSelect parece parametro inutil. jones 2002
Function GeracaoReqPedVenda_Le_Pedidos3(ByVal lComando As Long, ByVal gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda, iSelect As Integer) As Long
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao dos bloqueios

Dim lErro As Long
Dim viFilialEmpresa As Variant, vlPedidosDe As Variant, vlPedidosAte As Variant, vlClientesDe As Variant, vlClientesAte As Variant, viFilialPedidoDe As Variant, viFilialPedidoAte As Variant, vdtEmissaoDe As Variant, vdtEmissaoAte As Variant, vdtEntregaDe As Variant, vdtEntregaAte As Variant
Dim vdtDataNula As Variant

On Error GoTo Erro_GeracaoReqPedVenda_Le_Pedidos3

    If gobjGeracaoReqPedVenda.iExibeTodos = vbUnchecked Then

        If gobjGeracaoReqPedVenda.PedidoDe <> 0 Then
            vlPedidosDe = gobjGeracaoReqPedVenda.PedidoDe
            lErro = Comando_BindVarInt(lComando, vlPedidosDe)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68177
        End If

        If gobjGeracaoReqPedVenda.PedidoAte <> 0 Then
            vlPedidosAte = gobjGeracaoReqPedVenda.PedidoAte
            lErro = Comando_BindVarInt(lComando, vlPedidosAte)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68178
        End If

        If gobjGeracaoReqPedVenda.ClienteDe <> 0 Then
            vlClientesDe = gobjGeracaoReqPedVenda.ClienteDe
            lErro = Comando_BindVarInt(lComando, vlClientesDe)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68179
        End If

        If gobjGeracaoReqPedVenda.ClienteAte <> 0 Then
            vlClientesAte = gobjGeracaoReqPedVenda.ClienteAte
            lErro = Comando_BindVarInt(lComando, vlClientesAte)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68180
        End If

        If gobjGeracaoReqPedVenda.DataEmissaoDe <> DATA_NULA Then
            vdtEmissaoDe = gobjGeracaoReqPedVenda.DataEmissaoDe
            lErro = Comando_BindVarInt(lComando, vdtEmissaoDe)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68181
        End If

        If gobjGeracaoReqPedVenda.DataEmissaoAte <> DATA_NULA Then
            vdtEmissaoAte = gobjGeracaoReqPedVenda.DataEmissaoAte
            lErro = Comando_BindVarInt(lComando, vdtEmissaoAte)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68182
        End If
        
        If gobjGeracaoReqPedVenda.DataEntregaDe <> DATA_NULA Then
            vdtEntregaDe = gobjGeracaoReqPedVenda.DataEntregaDe
            lErro = Comando_BindVarInt(lComando, vdtEntregaDe)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68184
        End If

        If gobjGeracaoReqPedVenda.DataEntregaAte <> DATA_NULA Then
            vdtEntregaAte = gobjGeracaoReqPedVenda.DataEntregaAte
            lErro = Comando_BindVarInt(lComando, vdtEntregaAte)
            If (lErro <> AD_SQL_SUCESSO) Then Error 68185
        End If
    
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 68186
    
    GeracaoReqPedVenda_Le_Pedidos3 = SUCESSO

    Exit Function

Erro_GeracaoReqPedVenda_Le_Pedidos3:

    GeracaoReqPedVenda_Le_Pedidos3 = gErr

    Select Case gErr

        Case 68176 To 68186, 68315, 75401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_RC", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147657)

    End Select

    Exit Function

End Function

Function GeracaoReqPedVenda_Le_Pedidos4(ByVal lComando As Long, gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda, dtDataEntrega As Date, sNomeCliente As String, iFilialEmpresa As Integer, lCodigo As Long, lCliente As Long, iFilial As Integer, dtEmissaoPedido As Date, dValorTotal As Double) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objReqCompra As New ClassRequisicaoCompras
Dim objItemRC As New ClassItemReqCompras
Dim objPedVenda As New ClassPedidoDeVenda

On Error GoTo Erro_GeracaoReqPedVenda_Le_Pedidos4

    'Início Comentário Por Rafael Menezes em 12/09/2002
     '   Set gobjGeracaoReqPedVenda = New ClassGeracaoReqPedVenda
     '   Set gobjGeracaoReqPedVenda.colPedVenda = New Collection
    'Fim Comentário Por Rafael Menezes em 12/09/2002

    'Busca o primeiro Pedido de Venda no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68187

    If lErro = AD_SQL_SEM_DADOS Then gError 68188

    Do While lErro = AD_SQL_SUCESSO

        Set objPedVenda = New ClassPedidoDeVenda
        
        objPedVenda.iFilialEmpresa = iFilialEmpresa
        objPedVenda.lCodigo = lCodigo
        objPedVenda.lCliente = lCliente
        objPedVenda.iFilial = iFilial
        objPedVenda.dtDataEmissao = dtEmissaoPedido
        objPedVenda.dValorTotal = dValorTotal

        'Lê os itens do Pedido de Venda
        lErro = CF("PedidoDeVenda_Le_Itens", objPedVenda)
        If lErro <> SUCESSO Then gError 68137
        
        'adiciona o Pedido de Venda em gobjGeracaoReqPedVenda.colPedVenda
        gobjGeracaoReqPedVenda.colPedVenda.Add objPedVenda

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68189

    Loop

    GeracaoReqPedVenda_Le_Pedidos4 = SUCESSO

    Exit Function

Erro_GeracaoReqPedVenda_Le_Pedidos4:

    GeracaoReqPedVenda_Le_Pedidos4 = gErr

    Select Case gErr

        Case 68197, 68189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_RC", gErr)

        Case 68188 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147658)

    End Select

    Exit Function

End Function

Function FilialEmpresa_Lock(ByVal iCodFilial As Integer) As Long
'Lê e Faz "lockshared" na FilialEmpresa passada por iCodFilial

Dim lErro As Long
Dim iAlmox As Integer
Dim lComando As Long

On Error GoTo Erro_FilialEmpresa_Lock

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61726

    'Lê a Filial Empresa passada pelo Código em iCodFilial
    lErro = Comando_ExecutarPos(lComando, "SELECT AlmoxarifadoPadrao FROM FiliaisEmpresa WHERE FilialEmpresa = ?", 0, iAlmox, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 61723

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61724

    'Se não encontrou a FiliaEmpresa, Erro
    If lErro <> AD_SQL_SUCESSO Then Error 61725

    'Faz LockShared em FilialEmpresa
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61727

    'Fecha comando
    Call Comando_Fechar(lComando)

    FilialEmpresa_Lock = SUCESSO

    Exit Function

Erro_FilialEmpresa_Lock:

    FilialEmpresa_Lock = Err

    Select Case Err

        Case 61723, 61724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, iCodFilial)

        Case 61725 'FilialEmpresa não cadastrada

        Case 61726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIALEMPRESA", Err, iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147659)

    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Ccl_Lock(ByVal sCcl As String) As Long
'Lê e Faz "LockShared" do Ccl

Dim sCcl1 As String
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Ccl_Lock

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61728

    sCcl1 = String(STRING_CCL, 0)

    'Lê o Ccl passado
    lErro = Comando_ExecutarPos(lComando, "SELECT Ccl FROM Ccl WHERE Ccl = ? ", 0, sCcl1, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 61729

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61730

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61731

    'Faz LockShared no Ccl
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61738

    Call Comando_Fechar(lComando)

    Ccl_Lock = SUCESSO

    Exit Function

Erro_Ccl_Lock:

    Ccl_Lock = Err

    Select Case Err

        Case 61728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61729, 61730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, sCcl)

        Case 61731 'Ccl não cadastrado

        Case 61738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147660)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Requisitante_Lock(ByVal lCodRequisitante As Long) As Long
'Lê e faz LockShared no Requisitante

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_Requisitante_Lock

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61734

    'Tenta ler o Requisitante passado pelo Código
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM Requisitante WHERE Codigo = ?", 0, lCodigo, lCodRequisitante)
    If lErro <> AD_SQL_SUCESSO Then Error 61735

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61736

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61737

    'Faz "LockShared" no Requisitante
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61739

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Requisitante_Lock = SUCESSO

    Exit Function

Erro_Requisitante_Lock:

    Requisitante_Lock = Err

    Select Case Err

        Case 61734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61735, 61736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE", Err, lCodRequisitante)

        Case 61737 'Requisitante não cadastrado

        Case 61739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISITANTE", Err, lCodRequisitante)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147661)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Observacao_Lock(ByVal sObservacao As String, lObservacao As Long) As Long
'Faz "LockShared" em Observacao caso sObservacao já exista
'e insere nova Observação caso sObservacao não exista.
'Devolve em lObservacao o lNumInt gerado

Dim lErro As Long
Dim lNumInt As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Observacao_Lock

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 61720
    Next

    'Busca na tabela de Observacao a observacao da Requisição de Compras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumInt FROM Observacao WHERE Observacao = ?", 0, lNumInt, sObservacao)
    If lErro <> AD_SQL_SUCESSO Then Error 61748

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61749

    'Se nao encontrou
    If lErro = AD_SQL_SEM_DADOS Then

        'Gera código automatico
        lErro = CF("Observacao_Automatico", lNumInt)
        If lErro <> SUCESSO Then Error 61750

        'Insere observacao na tabela de Observacao
        lErro = Comando_Executar(alComando(1), "INSERT INTO Observacao(Observacao,NumInt)VALUES (?,?)", sObservacao, lNumInt)
        If lErro <> AD_SQL_SUCESSO Then Error 61751

        lObservacao = lNumInt

    Else

        'Faz "LockShared" Em Observacao
        lErro = Comando_LockShared(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 61752

        lObservacao = lNumInt

    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Observacao_Lock = SUCESSO

    Exit Function

Erro_Observacao_Lock:

    Observacao_Lock = Err

    Select Case Err

        Case 61720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61748, 61749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBSERVACAO", Err, sObservacao)

        Case 61750

        Case 61751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OBSERVACAO", Err, sObservacao)

        Case 61752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_OBSERVACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147662)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ContaContabil_Lock(ByVal sConta As String) As Long
'Lê e Faz Lock na Conta passada por sConta

Dim lErro As Long
Dim lComando As Long
Dim iTipoConta As Integer

On Error GoTo Erro_ContaContabil_Lock

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61759

    'Tenta ler a Conta Passada
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoConta FROM PlanoConta WHERE Conta =?", 0, iTipoConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 61760

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61761

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61762

    'Faz "lockShared" na Conta
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61763

    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    ContaContabil_Lock = SUCESSO

    Exit Function

Erro_ContaContabil_Lock:

    ContaContabil_Lock = Err

    Select Case Err

        Case 61759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61760, 61761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)

        Case 61762 'Conta não Cadastrada

        Case 61763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147663)

    End Select

    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposTributacaoCompras_Le(ByVal colTributacao As AdmColCodigoNome) As Long
'Lê Tipos de Tributação associadas a Compras

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim iCodigo As Integer

On Error GoTo Erro_TiposTributacaoCompras_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66581

    sDescricao = String(STRING_DESCRICAO_TIPOTRIBUTACAO, 0)

    'Lê os Tipos de Tributação associados a Compras
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao FROM TiposDeTributacaoMovto WHERE Compra = ?", iCodigo, sDescricao, TIPO_COMPRAS)
    If lErro <> AD_SQL_SUCESSO Then gError 66582

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66583

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66585

    Do While lErro = AD_SQL_SUCESSO

        'Adiciona na coleção
        Call colTributacao.Add(iCodigo, sDescricao)

        'Busca Próximo Tipo de Tributação
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66584

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposTributacaoCompras_Le = SUCESSO

    Exit Function

Erro_TiposTributacaoCompras_Le:

    TiposTributacaoCompras_Le = gErr

    Select Case gErr

        Case 66581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66582, 66583, 66584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO", gErr)

        Case 66585 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147664)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensReqCompra_Lock_Gravacao(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz lock nas tabelas associados ao item da Requisição Compras

Dim lErro As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim objProduto As New ClassProduto
Dim lComando As Long
Dim lObservacao As Long

On Error GoTo Erro_ItensReqCompra_Lock_Gravacao

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError (65521)

    'Para cada item da Requisição de Compras
    For Each objItemReqCompra In objRequisicaoCompra.colItens

        'Lê e Faz Lock do Produto
        objProduto.sCodigo = objItemReqCompra.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then gError (65522)

        'Se o Almoxarifado estiver preenchido
        If objItemReqCompra.iAlmoxarifado <> 0 Then

            'Lê e Faz Lock do Almoxarifado
            lErro = CF("Almoxarifado_Lock", objItemReqCompra.iAlmoxarifado, lComando)
            If lErro <> SUCESSO Then gError (65523)

        End If

        'Se o Ccl estiver preenchido
        If Len(Trim(objItemReqCompra.sCcl)) > 0 Then

            'Lê e Faz lock do Ccl
            lErro = CF("Ccl_Lock", objItemReqCompra.sCcl)
            If lErro <> SUCESSO And lErro <> 61731 Then gError (65524)
            If lErro = 61731 Then gError (65525)

        End If

        'Se a Conta Contábil estiver preenchida
        If Len(Trim(objItemReqCompra.sContaContabil)) > 0 Then

            'Lê e Faz Lock da Conta Contábil
            lErro = CF("ContaContabil_Lock", objItemReqCompra.sContaContabil)
            If lErro <> SUCESSO And lErro <> 61762 Then gError (65526)
            If lErro = 61762 Then gError (65527)

        End If

        'Se Fornecedor e Filial estiverem Preenchidos
        If objItemReqCompra.lFornecedor <> 0 And objItemReqCompra.iFilial <> 0 Then

            'Lê e Faz Lock em FilialFornecedor
            lErro = CF("FilialFornecedor_Lock", objItemReqCompra.lFornecedor, objItemReqCompra.iFilial)
            If lErro <> SUCESSO And lErro <> 42870 Then gError (65528)
            If lErro = 42870 Then gError (65529)

        End If

        'Se a Observação estiver Preenchida
        If Len(Trim(objItemReqCompra.sObservacao)) > 0 Then

            'Lê e Faz Lock em Observação
            lErro = CF("Observacao_Lock", objItemReqCompra.sObservacao, lObservacao)
            If lErro <> SUCESSO Then gError (65530)

            objItemReqCompra.lObservacao = lObservacao

        End If

    Next

    'Fecha comando
    Call Comando_Fechar(lComando)

    ItensReqCompra_Lock_Gravacao = SUCESSO

    Exit Function

Erro_ItensReqCompra_Lock_Gravacao:

    ItensReqCompra_Lock_Gravacao = gErr

    Select Case gErr

        Case 65521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65522, 65523, 65524, 65526, 65528, 65530

        Case 65525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemReqCompra.sCcl)

        Case 65527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objItemReqCompra.sContaContabil)

        Case 65529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objItemReqCompra.lFornecedor, objItemReqCompra.iFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147665)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicaoCompra_Lock_Gravacao(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz Lock nas tabelas Associadas a Requisição de Compras

Dim lErro As Long
Dim lObservacao As Long

On Error GoTo Erro_RequisicaoCompra_Lock_Gravacao

    'Se a FilialCompra tiver sido preenchida
    If objRequisicaoCompra.iFilialCompra <> 0 Then

        'Le e locka FilialEmpresa
        lErro = CF("FilialEmpresa_Lock", objRequisicaoCompra.iFilialCompra)
        If lErro <> SUCESSO And lErro <> 61725 Then gError (65508)
        If lErro = 61725 Then gError (65509)

    End If

    'Se o Ccl está preenchido
    If Len(Trim(objRequisicaoCompra.sCcl)) > 0 Then

        'Lê o locka Ccl
        lErro = CF("Ccl_Lock", objRequisicaoCompra.sCcl)
        If lErro <> SUCESSO And lErro <> 61731 Then gError (65510)
        If lErro = 61731 Then gError (65511)

    End If

    'Se o Requisitante estiver preenchido
    If objRequisicaoCompra.lRequisitante <> 0 Then

        'Lê e Faz Lock do Requisitante
        lErro = CF("Requisitante_Lock", objRequisicaoCompra.lRequisitante)
        If lErro <> SUCESSO And lErro <> 61737 Then gError (65512)
        If lErro = 61737 Then gError (65513)

    End If

    'Trata o Lock do Tipo Destino
    If objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA Then

        lErro = CF("FilialEmpresa_Lock", objRequisicaoCompra.iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 61725 Then gError (65514)

        'Se não encontrou FilialEmpresa, Erro
        If lErro = 61725 Then gError (65515)

    ElseIf objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

        lErro = CF("FilialFornecedor_Lock", objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino)
        If lErro <> SUCESSO And lErro <> 42870 Then gError (65518)

        'Se nao encontrou FilialFornecedor, Erro
        If lErro = 42870 Then gError (65519)

    End If

    'Verifica se observacao esta preenchida
    If Len(Trim(objRequisicaoCompra.sObservacao)) > 0 Then

        'Faz Lock e insere nova Observação
        lErro = CF("Observacao_Lock", objRequisicaoCompra.sObservacao, lObservacao)
        If lErro <> SUCESSO Then gError (65520)

        objRequisicaoCompra.lObservacao = lObservacao

    End If

    RequisicaoCompra_Lock_Gravacao = SUCESSO

    Exit Function

Erro_RequisicaoCompra_Lock_Gravacao:

    RequisicaoCompra_Lock_Gravacao = gErr

    Select Case gErr

        Case 65508, 65510, 65512, 65514, 65518, 65520

        Case 65509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRequisicaoCompra.iFilialCompra)

        Case 65511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objRequisicaoCompra.sCcl)

        Case 65513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisicaoCompra.lRequisitante)

        Case 65515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRequisicaoCompra.iFilialEmpresa)

        Case 65519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147666)

    End Select

    Exit Function

End Function

Function Cotacoes_Produto_Le(ByVal colGeracao As Collection, ByVal objProduto As ClassProduto, ByVal dQuantidade As Double, ByVal iTipoDestino As Integer, ByVal lDestino As Long, ByVal iFilialDestino As Integer, Optional ByVal lFornExclu As Long = 0, Optional ByVal iFilialExclu As Integer = 0) As Long
'Busca no BD os Cotacões cujo o Produto é o mesmo de sProduto
'respeitando os limites de quantidade

Dim lComando As Long
Dim lErro As Long
Dim objCotItemConc As ClassCotacaoItemConc
Dim tCotItemConc As typeCotacaoItemConcorrencia
Dim dQuantMais As Double, dQuantMenos As Double
Dim objCondPagto As New ClassCondicaoPagto
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCondPagto As Integer, iIndice As Integer
Dim lFornecedor As Long, iFilialForn As Integer
Dim dtData As Date, bAchou As Boolean
Dim colCondPagtoForn As Collection
Dim lFornAtual As Long
Dim iFilialAtual As Long
Dim dtDataAtual As Date
Dim colFornec As New AdmCollCodigoNome
Dim colFiliaisForn As New Collection
Dim objCodNome As AdmlCodigoNome
Dim iPosicao As Integer
Dim dQuantComprarMax As Double
Dim colMoedas As New Collection
Dim bAchou2 As Boolean
Dim iIndice2 As Integer

On Error GoTo Erro_Cotacoes_Produto_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 63780

    'Se o Produto Considera a quantidade de Cotação Anterior
    If objProduto.iConsideraQuantCotAnt = PRODUTO_CONSIDERA_QUANT_COTACAO_ANTERIOR Then

        'Calcula o Intervalo de quantidade a considerar
        dQuantMenos = dQuantidade - dQuantidade * objProduto.dPercentMenosQuantCotAnt
        dQuantMais = dQuantidade + dQuantidade * objProduto.dPercentMaisQuantCotAnt
        
        'Se nenhum tipo de Destino foi selecionado
        If iTipoDestino = TIPO_DESTINO_AUSENTE Then
            
            'Se a filial do fornecedor
            If lFornExclu <> 0 And iFilialExclu <> 0 Then
                'Pesquisa no Banco, as cotações mais recentes que se enquadrem no intervalo de quantidade e tiop de destino indicado e Fornecedor e Filialforn exclusivos
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao,Cotacao " & _
                " WHERE Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.Quantidade >= ? AND CotacaoProduto.Quantidade <= ? AND CotacaoProduto.UM = ? And PedidoCotacao.FilialEmpresa = ? AND PedidoCotacao.Fornecedor = ? AND PedidoCotacao.Filial = ? AND (PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, dQuantMenos, dQuantMais, objProduto.sSiglaUMCompra, giFilialEmpresa, lFornExclu, iFilialExclu, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
                
            Else
                
                'Pesquisa no Banco, as cotações mais recentes que se enquadrem no intervalo de quantidade e tiop de destino indicado
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao,Cotacao " & _
                " WHERE Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.Quantidade >= ? AND CotacaoProduto.Quantidade <= ? AND CotacaoProduto.UM = ? And PedidoCotacao.FilialEmpresa = ? AND (PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, dQuantMenos, dQuantMais, objProduto.sSiglaUMCompra, giFilialEmpresa, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            End If
        Else



            If lFornExclu <> 0 And iFilialExclu <> 0 Then

                'Pesquisa no Banco, as cotações mais recentes que se enquadrem no intervalo de quantidade e tiop de destino indicado e Fornecedor e Filialforn exclusivos
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao,Cotacao" & _
                " WHERE Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.Quantidade >= ? AND CotacaoProduto.Quantidade <= ? AND CotacaoProduto.UM = ? And PedidoCotacao.FilialEmpresa = ? AND PedidoCotacao.Fornecedor = ? AND PedidoCotacao.Filial = ? AND ((Cotacao.TipoDestino = ?) OR (Cotacao.TipoDestino= ? AND Cotacao.FornCliDestino= ? AND Cotacao.FilialDestino= ?)) AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, dQuantMenos, dQuantMais, objProduto.sSiglaUMCompra, giFilialEmpresa, lFornExclu, iFilialExclu, TIPO_DESTINO_AUSENTE, iTipoDestino, lDestino, iFilialDestino, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
                
            Else
                
                'Pesquisa no Banco, as cotações mais recentes que se enquadrem no intervalo de quantidade e tiop de destino indicado
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao, Cotacao" & _
                " WHERE  Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.Quantidade >= ? AND CotacaoProduto.Quantidade <= ? AND CotacaoProduto.UM = ? And PedidoCotacao.FilialEmpresa = ? AND ((Cotacao.TipoDestino = ?) OR (Cotacao.TipoDestino= ? AND Cotacao.FornCliDestino= ? AND Cotacao.FilialDestino= ?)) AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, dQuantMenos, dQuantMais, objProduto.sSiglaUMCompra, giFilialEmpresa, TIPO_DESTINO_AUSENTE, iTipoDestino, lDestino, iFilialDestino, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            End If
       
        End If
    Else
        If iTipoDestino = TIPO_DESTINO_AUSENTE Then
            If lFornExclu <> 0 And iFilialExclu <> 0 Then
                'Pesquisa no Banco, as cotações mais recentes vinculadas a concorrências
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao, Cotacao" & _
                " WHERE CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.UM = ? AND PedidoCotacao.FilialEmpresa = ? AND PedidoCotacao.Fornecedor = ? AND PedidoCotacao.Filial = ?  AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, objProduto.sSiglaUMCompra, giFilialEmpresa, lFornExclu, iFilialExclu, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            Else
                'Pesquisa no Banco, as cotações mais recentes vinculadas a concorrências
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS,ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao, Cotacao" & _
                " WHERE CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND  CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.UM = ? AND PedidoCotacao.FilialEmpresa = ?  AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, objProduto.sSiglaUMCompra, giFilialEmpresa, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            End If
        Else
            If lFornExclu <> 0 And iFilialExclu <> 0 Then
                'Pesquisa no Banco, as cotações mais recentes vinculadas a concorrências
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS, ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao, Cotacao" & _
                " WHERE CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND  CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.UM = ? AND PedidoCotacao.FilialEmpresa = ? AND PedidoCotacao.Fornecedor = ? AND PedidoCotacao.Filial = ? AND ((Cotacao.TipoDestino = ?) OR (Cotacao.TipoDestino= ? AND Cotacao.FornCliDestino= ? AND Cotacao.FilialDestino= ?)) AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, objProduto.sSiglaUMCompra, giFilialEmpresa, lFornExclu, iFilialExclu, TIPO_DESTINO_AUSENTE, iTipoDestino, lDestino, iFilialDestino, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            Else
                'Pesquisa no Banco, as cotações mais recentes vinculadas a concorrências
                lErro = Comando_Executar(lComando, "SELECT ItensCotacao.AliquotaIPI, ItensCotacao.AliquotaICMS,ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Codigo, PedidoCotacao.DataValidade, ItensCotacao.PrazoEntrega, ItensCotacao.QuantEntrega, CotacaoProduto.Quantidade, ItensCotacao.NumIntDoc, ItensCotacao.DataReferencia, PedidoCotacao.Data, ItensCotacao.Moeda, ItensCotacao.Taxa FROM CotacaoProduto, ItensCotacao, PedidoCotacao, ItemPedCotacao, Cotacao" & _
                " WHERE CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND  CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto AND ItemPedCotacao.PedidoCotacao = PedidoCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao = ItemPedCotacao.NumIntDoc AND CotacaoProduto.Produto = ? AND CotacaoProduto.UM = ? AND PedidoCotacao.FilialEmpresa = ? AND ((Cotacao.TipoDestino = ?) OR (Cotacao.TipoDestino= ? AND Cotacao.FornCliDestino= ? AND Cotacao.FilialDestino= ?)) AND(PedidoCotacao.DataValidade = ? OR PedidoCotacao.DataValidade >= ?) " & _
                " ORDER BY PedidoCotacao.Fornecedor, PedidoCotacao.Filial, PedidoCotacao.Data DESC, ItensCotacao.CondPagto, PedidoCotacao.Codigo DESC, ItensCotacao.PrecoUnitario", _
                tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
                tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, objProduto.sCodigo, objProduto.sSiglaUMCompra, giFilialEmpresa, TIPO_DESTINO_AUSENTE, iTipoDestino, lDestino, iFilialDestino, DATA_NULA, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 63819
            End If
        End If
    
    End If

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63781

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 63822
        
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
        
        Set colCondPagtoForn = New Collection
        Set colMoedas = New Collection
        lFornAtual = lFornecedor
        iFilialAtual = iFilialForn
        dtDataAtual = dtData
    
        'Para as ultimas cotações da filial fornecedor selecionada
        Do While (lErro = AD_SQL_SUCESSO) And (lFornAtual = lFornecedor And iFilialAtual = iFilialForn And dtDataAtual = dtData)

            bAchou = False
            
            'verifica se já tem uma cotação feita para essa cond Pagto
            For iIndice = 1 To colCondPagtoForn.Count
                If iCondPagto = colCondPagtoForn(iIndice) Then
                    For iIndice2 = 1 To colMoedas.Count
                        If tCotItemConc.iMoeda = colMoedas(iIndice) Then
                            bAchou = True
                            Exit For
                        Else
                            bAchou = False
                        End If
                    Next
                Else
                    bAchou = False
                End If
            Next
                
            'Se não tem
            If Not bAchou Then
                
                'recolhe a cotação
                colCondPagtoForn.Add iCondPagto
                colMoedas.Add tCotItemConc.iMoeda
                
                Set objCotItemConc = New ClassCotacaoItemConc
        
                'Lê a Condição de pagamento
                objCondPagto.iCodigo = iCondPagto
                lErro = CF("CondicaoPagto_Le", objCondPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 63782
        
                'se nao encontrou --->erro
                If lErro = 19205 Then gError 63855
        
                'Coloca como default, quantidade a comprar igual a quant. max. à comprar
               
                objCotItemConc.sCondPagto = objCondPagto.iCodigo & SEPARADOR & objCondPagto.sDescReduzida
                objCotItemConc.lItemCotacao = tCotItemConc.lItemCotacao
                objCotItemConc.dPrecoUnitario = tCotItemConc.dPrecoUnitario
                objCotItemConc.dtDataPedidoCotacao = dtData
                
                'Quantidade máxima à Comprar
                objCotItemConc.dQuantidadeComprar = dQuantidade
        
                If lFornecedor <> 0 And iFilialForn <> 0 Then
        
                    Call Busca_Na_Colecao(colFornec, lFornecedor, iPosicao)
                    
                    If iPosicao = 0 Then
        
                        'Lê o Fornecedor
                        objFornecedor.lCodigo = lFornecedor
                        lErro = CF("Fornecedor_Le", objFornecedor)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError 63783
            
                        'Se não encontrou, Erro
                        If lErro = 12729 Then gError 63856
                        
                        Set objCodNome = New AdmlCodigoNome
                                
                        objCodNome.lCodigo = lFornecedor
                        objCodNome.sNome = objFornecedor.sNomeReduzido
                        
                        colFornec.Add objCodNome.lCodigo, objCodNome.sNome
        
                    Else
                        Set objCodNome = colFornec(iPosicao)
                    End If
                    
                    objCotItemConc.sFornecedor = objCodNome.sNome
        
                    Call Busca_FilialForn(colFiliaisForn, lFornecedor, iFilialForn, iPosicao)
        
                    If iPosicao = 0 Then
                        'Lê a Filial do Fornecedor
                        objFilialFornecedor.iCodFilial = iFilialForn
                        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
                        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError 63785
            
                        'Se não encontrou, Erro
                        If lErro = 12929 Then gError 63779
                        
                        colFiliaisForn.Add objFilialFornecedor
                    Else
                        Set objFilialFornecedor = colFiliaisForn(iPosicao)
                    End If
                    
                    objCotItemConc.sFilial = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
                    
                End If
        
                objCotItemConc.lPedCotacao = tCotItemConc.lPedCotacao
                objCotItemConc.dtDataValidade = tCotItemConc.dtDataValidade
                objCotItemConc.iPrazoEntrega = tCotItemConc.iPrazoEntrega
                objCotItemConc.dQuantEntrega = tCotItemConc.dQuantEntrega
                objCotItemConc.dAliquotaIPI = tCotItemConc.dAliquotaIPI
                objCotItemConc.dAliquotaICMS = tCotItemConc.dCreditoICMS
                objCotItemConc.sUMCompra = objProduto.sSiglaUMCompra
                objCotItemConc.dQuantCotada = dQuantComprarMax
                objCotItemConc.dPrecoAjustado = objCotItemConc.dPrecoUnitario
                objCotItemConc.iSelecionada = MARCADO
                objCotItemConc.iMoeda = tCotItemConc.iMoeda
                objCotItemConc.dTaxa = tCotItemConc.dTaxa
                If tCotItemConc.dtDataEntrega <> DATA_NULA Then
                    objCotItemConc.dtDataEntrega = DateAdd("d", tCotItemConc.iPrazoEntrega, tCotItemConc.dtDataEntrega)
                End If
        
                'Se estiver Adiciona na coleção
                colGeracao.Add objCotItemConc
        
            End If
            'Busca próxima Cotação
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63784
        Loop
        
        'Elimita as cotações anteriores do Forn/filial
        If (lFornAtual = lFornecedor And iFilialAtual = iFilialForn And dtDataAtual <> dtData) Then
            Do While lErro = AD_SQL_SUCESSO And (lFornAtual = lFornecedor And iFilialAtual = iFilialForn)
                'Busca próxima Cotação
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63784
            Loop
        End If
        
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cotacoes_Produto_Le = SUCESSO

    Exit Function

Erro_Cotacoes_Produto_Le:

    Cotacoes_Produto_Le = gErr

    Select Case gErr

        Case 63780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 63819, 63781, 63784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case 63822, 63782, 63783, 63785, 70487

        Case 63855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondPagto.iCodigo)

        Case 63856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 63779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 70488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147667)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoItemConcorrencia_Le_ItemConc(lNumItemCotacao As Long, ByVal lNumItemConc As Long, ByVal lNumIntDoc As Long) As Long
'Lê o NumIntDoc da Cotação Item Concorrência partir do ItemConcorrência e ItemCotação

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_CotacaoItemConcorrencia_Le_ItemConc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67001
    
     'Verifica se a CotaçãoItemConcorrência existe
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM CotacaoItemConcorrencia WHERE ItemConcorrencia = ? AND ItemCotacao = ?", lNumIntDoc, lNumItemConc, lNumItemCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 67002

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67003

    'CotaçãoItemConcorrência não está cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 67004
    
    lErro = Comando_Fechar(lComando)

    CotacaoItemConcorrencia_Le_ItemConc = SUCESSO

    Exit Function

Erro_CotacaoItemConcorrencia_Le_ItemConc:

    CotacaoItemConcorrencia_Le_ItemConc = gErr

    Select Case gErr

        Case 67001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67002, 67003
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 67004 'Cotação Item Concorrência não cadastrada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147668)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function CotacaoProduto_Le(ByVal objCotacao As ClassCotacao) As Long
'Lê Cotações do Produto relacionadas a uma cotação a partir do NumIntDoc passado em objCotacao

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoProduto As typeCotacaoProduto
Dim objCotacaoProduto As ClassCotacaoProduto

On Error GoTo Erro_CotacaoProduto_Le

    'Abertura de Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67142
    
    tCotacaoProduto.sProduto = String(STRING_PRODUTO, 0)
    tCotacaoProduto.sUM = String(STRING_UM_SIGLA, 0)
    
    'Lê Cotacao Produto a partir do NumIntDoc da Cotação
    lErro = Comando_Executar(lComando, "SELECT  NumIntDoc, Cotacao, Produto, Quantidade, UM, Fornecedor, Filial FROM CotacaoProduto WHERE Cotacao = ? ", tCotacaoProduto.lNumIntDoc, tCotacaoProduto.lCotacao, tCotacaoProduto.sProduto, tCotacaoProduto.dQuantidade, tCotacaoProduto.sUM, tCotacaoProduto.lFornecedor, tCotacaoProduto.iFilial, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67143

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67144

    'Se não encontrou, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67145

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
    
        Set objCotacaoProduto = New ClassCotacaoProduto
        
        'Armazena as informações lidas do BD
        objCotacaoProduto.lNumIntDoc = tCotacaoProduto.lNumIntDoc
        objCotacaoProduto.lCotacao = tCotacaoProduto.lCotacao
        objCotacaoProduto.sProduto = tCotacaoProduto.sProduto
        objCotacaoProduto.dQuantidade = tCotacaoProduto.dQuantidade
        objCotacaoProduto.sUM = tCotacaoProduto.sUM
        objCotacaoProduto.lFornecedor = tCotacaoProduto.lFornecedor
        objCotacaoProduto.iFilial = tCotacaoProduto.iFilial
        
        'Adiciona na coleção
        objCotacao.colCotacaoProduto.Add objCotacaoProduto
        
        'Busca próxima Cotação Produto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67146
    
    Loop
    
    'Fechamento do comando
    Call Comando_Fechar(lComando)

    CotacaoProduto_Le = SUCESSO

    Exit Function

Erro_CotacaoProduto_Le:

    CotacaoProduto_Le = gErr

    Select Case gErr

        Case 67142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67143, 67144, 67146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)

        Case 67145 'Não existe na tabela Cotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147669)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoProdutoItemRC_Le_Todas(ByVal objCotacao As ClassCotacao) As Long
'ItensRC relacionadas a uma cotação Produto a partir do NumIntDoc das Cotações Produtos passadas
'na coleção de objCotacao

Dim lErro As Long
Dim lComando As Long
Dim lCotacaoProduto As Long
Dim lItemReqCompra As Long
Dim dQuantCotar As Double
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemRC As ClassItemReqCompras

On Error GoTo Erro_CotacaoProdutoItemRC_Le_Todas

    'Abertura de Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67147
        
    'Para cada Cotação produto da Cotação passada
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto
    
        Set objCotacaoProduto.colItemReqCompras = New Collection
        
        'Lê Cotacao Produto Item RC a partir do NumIntDoc da Cotação Produto
        lErro = Comando_Executar(lComando, "SELECT CotacaoProduto, ItemReqCompra, QuantidadeCotar FROM CotacaoProdutoItemRC WHERE CotacaoProduto = ? ", lCotacaoProduto, lItemReqCompra, dQuantCotar, objCotacaoProduto.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 67148

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67149

        'Enquanto encontrar
        Do While lErro = AD_SQL_SUCESSO
            
            Set objItemRC = New ClassItemReqCompras
        
            'Armazena as informações lidas do BD
            objItemRC.lNumIntDoc = lItemReqCompra
            objItemRC.dQuantCotar = dQuantCotar
            
            'Adiciona na coleção
            objCotacaoProduto.colItemReqCompras.Add objItemRC
        
            'Busca próxima Cotação Produto
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67150
    
        Loop
    
    Next
    
    'Fechamento do comando
    Call Comando_Fechar(lComando)

    CotacaoProdutoItemRC_Le_Todas = SUCESSO

    Exit Function

Erro_CotacaoProdutoItemRC_Le_Todas:

    CotacaoProdutoItemRC_Le_Todas = gErr

    Select Case gErr

        Case 67147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67148, 67149, 67150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147670)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function



Function GeracaoPedCompraCot_Le_GeracaoPedCotacao(ByVal objGeracaoPedidoCompraCot As ClassGeracaoPedCompraCot) As Long
'Lê as Cotações selecionadas a partir das opções preenchidas

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim lCodigoCot As Long, sDescricaoCot As String, lNumIntDocCot As Long, dtDataCot As Date
Dim dTaxaFinanceiraCot As Double, iTipoDestinoCot As Integer, lFornCliDestinoCot As Long, iFilialDestinoCot As Integer

On Error GoTo Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67196

    'Montar o SELECT dinamicamente
    Call GeracaoPedCompraCot_Le_GeracaoPedCotacao1(objGeracaoPedidoCompraCot, sSelect)

    sDescricaoCot = String(STRING_DESCRICAO_CAMPO, 0)

    'Preparar parte fixa do SELECT
    lErro = GeracaoPedCompraCot_Le_GeracaoPedCotacao2(lComando, sSelect, lCodigoCot, sDescricaoCot, lNumIntDocCot, dtDataCot, dTaxaFinanceiraCot, iTipoDestinoCot, lFornCliDestinoCot, iFilialDestinoCot)
    If lErro <> SUCESSO Then gError 67197

    'Complementa a passagem dos parâmetros que variam de acordo com a selecao do usuario
    'e executa o SELECT
    lErro = GeracaoPedCompraCot_Le_GeracaoPedCotacao3(lComando, objGeracaoPedidoCompraCot)
    If lErro <> SUCESSO Then gError 67198

    'Processar os registros lidos
    lErro = GeracaoPedCompraCot_Le_GeracaoPedCotacao4(lComando, objGeracaoPedidoCompraCot, lCodigoCot, sDescricaoCot, lNumIntDocCot, dtDataCot, dTaxaFinanceiraCot, iTipoDestinoCot, lFornCliDestinoCot, iFilialDestinoCot)
    If lErro <> SUCESSO And lErro <> 67201 Then gError 67199
    If lErro = 67201 Then gError 67200

    Call Comando_Fechar(lComando)

    GeracaoPedCompraCot_Le_GeracaoPedCotacao = SUCESSO

    Exit Function

Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao:

    GeracaoPedCompraCot_Le_GeracaoPedCotacao = gErr

    Select Case gErr

        Case 67196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67197, 67198, 67199
        
        Case 67200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_COTACAO_SELECAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147671)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Sub GeracaoPedCompraCot_Le_GeracaoPedCotacao1(ByVal objGeracaoPedidoCompraCot As ClassGeracaoPedCompraCot, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String
                
    sFields = "Cotacao.Codigo, Cotacao.Descricao, Cotacao.NumIntDoc, Cotacao.Data, Cotacao.TipoDestino, Cotacao.FornCliDestino, Cotacao.FilialDestino"
    sFrom = " FROM Cotacao, Comprador, CotacaoProduto, ItemPedCotacao, CotacaoProdutoItemRC"
    sWhere = " WHERE Comprador.CodUsuario = ? AND Cotacao.Comprador = Comprador.codigo AND Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.NumIntDoc = ItemPedCotacao.CotacaoProduto  AND CotacaoProdutoItemRC.CotacaoProduto = CotacaoProduto.NumIntDoc"

    If giFilialEmpresa <> EMPRESA_TODA Then sWhere = sWhere & " AND Cotacao.FilialEmpresa = ?"

    If objGeracaoPedidoCompraCot.dtDataDe <> DATA_NULA Then sWhere = sWhere & " AND Cotacao.Data >= ?"

    If objGeracaoPedidoCompraCot.dtDataAte <> DATA_NULA Then sWhere = sWhere & " AND Cotacao.Data <= ?"

    If objGeracaoPedidoCompraCot.lCodigoDe <> 0 Then sWhere = sWhere & " AND Cotacao.Codigo >= ?"

    If objGeracaoPedidoCompraCot.lCodigoAte <> 0 Then sWhere = sWhere & " AND Cotacao.Codigo <= ?"

    If objGeracaoPedidoCompraCot.iSelecionaDestino <> 0 Then

        sWhere = sWhere & " AND Cotacao.TipoDestino = ?"

        If objGeracaoPedidoCompraCot.lFornCliDestino <> 0 Then sWhere = sWhere & " AND Cotacao.FornCliDestino = ? "
        If objGeracaoPedidoCompraCot.iFilialDestino <> 0 Then sWhere = sWhere & " AND Cotacao.FilialDestino = ? "

    End If

    sSelect = "SELECT DISTINCT " & sFields & sFrom & sWhere

End Sub

Function GeracaoPedCompraCot_Le_GeracaoPedCotacao2(ByVal lComando As Long, ByVal sSelect As String, vlCodigoCot As Variant, vsDescricaoCot As Variant, vlNumIntDocCot As Variant, vdtDataCot As Variant, vdTaxaFinanceiraCot As Variant, viTipoDestinoCot As Variant, vlFornCliDestinoCot As Variant, viFilialDestinoCot As Variant) As Long
'Trata parâmetros que serão lidos no SELECT (PARTE FIXA)

Dim lErro As Long
Dim ret As Integer

On Error GoTo Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 67202

    ret = Comando_BindVarInt(lComando, vlCodigoCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67203

    ret = Comando_BindVarInt(lComando, vsDescricaoCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67204

    ret = Comando_BindVarInt(lComando, vlNumIntDocCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67205

    ret = Comando_BindVarInt(lComando, vdtDataCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67206

    ret = Comando_BindVarInt(lComando, viTipoDestinoCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67208

    ret = Comando_BindVarInt(lComando, vlFornCliDestinoCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67209

    ret = Comando_BindVarInt(lComando, viFilialDestinoCot)
    If (ret <> AD_SQL_SUCESSO) Then gError 67210

    GeracaoPedCompraCot_Le_GeracaoPedCotacao2 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao2:

    GeracaoPedCompraCot_Le_GeracaoPedCotacao2 = gErr

    Select Case gErr

        Case 67202 To 67210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147672)

    End Select

    Exit Function

End Function

Function GeracaoPedCompraCot_Le_GeracaoPedCotacao3(ByVal lComando As Long, ByVal objGeracaoPedCompra As ClassGeracaoPedCompraCot) As Long
'Trata a parte dinâmica do SELECT

Dim lErro As Long
Dim vsComprador As Variant, viFilialEmpresa As Variant, vdtDataDe As Variant, vdtDataAte As Variant, vlCodigoDe As Variant, vlCodigoAte As Variant
Dim viTipoDestino As Variant, vlFornCliDestino As Variant, viFilialDestino As Variant

On Error GoTo Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao3

    vsComprador = gsUsuario
    lErro = Comando_BindVarInt(lComando, vsComprador)
    If (lErro <> AD_SQL_SUCESSO) Then gError 67211

    If (giFilialEmpresa <> EMPRESA_TODA) Then
        viFilialEmpresa = giFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67212
    End If

    If objGeracaoPedCompra.dtDataDe <> DATA_NULA Then
        vdtDataDe = objGeracaoPedCompra.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67213
    End If

    If objGeracaoPedCompra.dtDataAte <> DATA_NULA Then
        vdtDataAte = objGeracaoPedCompra.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67214
    End If

    If objGeracaoPedCompra.lCodigoDe <> 0 Then
        vlCodigoDe = objGeracaoPedCompra.lCodigoDe
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67215
    End If

    If objGeracaoPedCompra.lCodigoAte <> 0 Then
        vlCodigoAte = objGeracaoPedCompra.lCodigoAte
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67216
    End If

    If objGeracaoPedCompra.iSelecionaDestino <> 0 Then

        viTipoDestino = objGeracaoPedCompra.iTipoDestino
        lErro = Comando_BindVarInt(lComando, viTipoDestino)
        If (lErro <> AD_SQL_SUCESSO) Then gError 67217

        If objGeracaoPedCompra.lFornCliDestino <> 0 Then
            vlFornCliDestino = objGeracaoPedCompra.lFornCliDestino
            lErro = Comando_BindVarInt(lComando, vlFornCliDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 67218
        End If

        If objGeracaoPedCompra.iFilialDestino <> 0 Then
            viFilialDestino = objGeracaoPedCompra.iFilialDestino
            lErro = Comando_BindVarInt(lComando, viFilialDestino)
            If (lErro <> AD_SQL_SUCESSO) Then gError 67219
        End If

    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 67220

    GeracaoPedCompraCot_Le_GeracaoPedCotacao3 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao3:

    GeracaoPedCompraCot_Le_GeracaoPedCotacao3 = gErr

    Select Case gErr

        Case 67211 To 67220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147673)

    End Select

    Exit Function

End Function

Function GeracaoPedCompraCot_Le_GeracaoPedCotacao4(ByVal lComando As Long, ByVal objGeracaoPedCompraCot As ClassGeracaoPedCompraCot, lCodigoCot As Long, sDescricaoCot As String, lNumIntDocCot As Long, dtDataCot As Date, dTaxaFinanceiraCot As Double, iTipoDestinoCot As Integer, lFornCliDestinoCot As Long, iFilialDestinoCot As Integer) As Long
'Executa o select e lê os campos

Dim lErro As Long
Dim objCotacao As ClassCotacao
Dim bAdiciona As Boolean
Dim iNumItensRC As Integer

On Error GoTo Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67221

    'Se não encontrou nenhuma Cotação, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67201

    'Enquanto encontrar registros
    Do While lErro = AD_SQL_SUCESSO

        Set objCotacao = New ClassCotacao

        objCotacao.lCodigo = lCodigoCot
        objCotacao.sDescricao = sDescricaoCot
        objCotacao.lNumIntDoc = lNumIntDocCot
        objCotacao.dtData = dtDataCot
        objCotacao.iTipoDestino = iTipoDestinoCot
        objCotacao.lFornCliDestino = lFornCliDestinoCot
        objCotacao.iFilialDestino = iFilialDestinoCot
        objCotacao.iFilialEmpresa = giFilialEmpresa

        bAdiciona = True
        'Se o Tipo Destino for ausente
        If iTipoDestinoCot = TIPO_DESTINO_AUSENTE Then
            
            lErro = Cotacao_Le_NumeroItensRC_Associados(objCotacao, iNumItensRC)
            If lErro <> SUCESSO Then Error 222222
            
            If iNumItensRC = 0 Then bAdiciona = False
            
        End If
        
        If bAdiciona Then
            'Adiciona Cotação na coleção de Cotações
            objGeracaoPedCompraCot.colCotacao.Add objCotacao
        End If
        
        'Busca próxima Cotação
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67222

    Loop

    GeracaoPedCompraCot_Le_GeracaoPedCotacao4 = SUCESSO

    Exit Function

Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao4:

    GeracaoPedCompraCot_Le_GeracaoPedCotacao4 = gErr

    Select Case gErr

        Case 67221, 67222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case 67201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147674)

    End Select

    Exit Function

End Function


Function ItemReqCompra_Le_Produto(ByVal objItemReqCompra As ClassItemReqCompras) As Long
'Lê o NumIntDoc do ItemRC a partir do NumIntDoc da Requisição o do Produto de ItemRC

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemReqCompra_Le_Produto

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66986

    'Lê na tabela de ItemReqCompra o NumIntDoc do ItemRC
    lErro = Comando_Executar(lComando, "SELECT ItensReqCompra.NumIntDoc FROM ItensReqCompra, RequisicaoCompra WHERE ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.Codigo = ? AND Produto = ?", lNumIntDoc, objItemReqCompra.lReqCompra, objItemReqCompra.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 66987

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66988

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66989

    'Guarda o número interno do ItemRC
    objItemReqCompra.lNumIntDoc = lNumIntDoc
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemReqCompra_Le_Produto = SUCESSO

    Exit Function

Erro_ItemReqCompra_Le_Produto:

    ItemReqCompra_Le_Produto = gErr

    Select Case gErr

        Case 66986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66987, 66988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 66989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMREQCOMPRA_NAO_CADASTRADO", gErr, objItemReqCompra.sProduto, objItemReqCompra.lReqCompra)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147675)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cotacao_Le_Tudo(ByVal objCotacao As ClassCotacao) As Long
'Lê a Cotação passada pelo NumIntDoc e coleções associadas
'(Cotação Produto e Cotação Produto ItemRC)

Dim lErro As Long

On Error GoTo Erro_Cotacao_Le_Tudo
    
    'Lê os dados da cotação
    lErro = CF("Cotacao_Le", objCotacao)
    If lErro <> SUCESSO Then gError 67139
    
    Set objCotacao.colCotacaoProduto = New Collection
    
    'Lê as Cotações Produto relacionadas a Cotação
    lErro = CF("CotacaoProduto_Le", objCotacao)
    If lErro <> SUCESSO Then gError 67140
    
    'Lê CotaçãoProdutoItemRC relacionadas a Cotação
    lErro = CF("CotacaoProdutoItemRC_Le_Todas", objCotacao)
    If lErro <> SUCESSO Then gError 67141
    
    Cotacao_Le_Tudo = SUCESSO
    
    Exit Function
    
Erro_Cotacao_Le_Tudo:

    Cotacao_Le_Tudo = gErr
    
    Select Case gErr
        
        Case 67139, 67140, 67141
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147676)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutoFilial_Le_PontoPedido(ByVal objGeracaoPtoPedido As ClassGeracaoPtoPedido) As Long

Dim lErro As Long
Dim sSelect As String
Dim tProduto As typeProduto
Dim tProdutoFilial As typeProdutoFilial
Dim tEstoqueProduto As typeEstoqueProduto
Dim tAlmoxarifado As typeAlmoxarifado
Dim lComando As Long

On Error GoTo Erro_ProdutoFilial_Le_PontoPedido

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67417

    tProduto.sCodigo = String(STRING_PRODUTO, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tEstoqueProduto.sContaContabil = String(STRING_CONTA, 0)
    tAlmoxarifado.sContaContabil = String(STRING_CONTA, 0)
    
    'Montar o SELECT dinamicamente
    Call ProdutoFilial_Le_PontoPedido1(objGeracaoPtoPedido, sSelect)

    'Preparar parte fixa do SELECT
    lErro = ProdutoFilial_Le_PontoPedido2(lComando, sSelect, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, tProdutoFilial.dPontoPedido, tProdutoFilial.dEstoqueSeguranca, tProdutoFilial.dConsumoMedio, tProdutoFilial.iIntRessup, tProdutoFilial.dLoteEconomico, tProdutoFilial.iAlmoxarifado, tEstoqueProduto.sContaContabil, tProdutoFilial.lFornecedor, tProdutoFilial.iFilialForn, tProdutoFilial.iFilialEmpresa, tAlmoxarifado.sContaContabil)
    If lErro <> SUCESSO Then gError 67418

    'Complementa a passagem dos parâmetros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das Requisições
    lErro = ProdutoFilial_Le_PontoPedido3(lComando, objGeracaoPtoPedido)
    If lErro <> SUCESSO Then gError 67419

    'Processar os registros lidos
    lErro = ProdutoFilial_Le_PontoPedido4(lComando, objGeracaoPtoPedido, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, tProdutoFilial.dPontoPedido, tProdutoFilial.dEstoqueSeguranca, tProdutoFilial.dConsumoMedio, tProdutoFilial.iIntRessup, tProdutoFilial.dLoteEconomico, tProdutoFilial.iAlmoxarifado, tEstoqueProduto.sContaContabil, tProdutoFilial.lFornecedor, tProdutoFilial.iFilialForn, tProdutoFilial.iFilialEmpresa, tAlmoxarifado.sContaContabil)
    If lErro <> SUCESSO And lErro <> 67442 Then gError 67420

    'se não encontrou - - - - > Erro
    If lErro = 67442 Then gError 67421

    Call Comando_Fechar(lComando)

    ProdutoFilial_Le_PontoPedido = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le_PontoPedido:

    ProdutoFilial_Le_PontoPedido = gErr

    Select Case gErr

        Case 67417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67418, 67419, 67420 'Tratados Nas rotinas chmadas

        Case 67421 'Sem Dados
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PRODUTO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147677)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub ProdutoFilial_Le_PontoPedido1(ByVal objGeracaoPtoPedido As ClassGeracaoPtoPedido, sSelect As String)
'Monta Select Dinamicamente

Dim sFields As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim objTipoProduto As New ClassTipoDeProduto
Dim iIndice As Integer

    sFields = "Produtos.Codigo, Produtos.Descricao, Produtos.SiglaUMEstoque, ProdutosFilial.PontoPedido, ProdutosFilial.EstoqueSeguranca, ProdutosFilial.ConsumoMedio, ProdutosFilial.IntRessup, ProdutosFilial.LoteEconomico, ProdutosFilial.Almoxarifado, EstoqueProduto.ContaContabil, ProdutosFilial.Fornecedor, ProdutosFilial.FilialForn, ProdutosFilial.FilialEmpresa, Almoxarifado.ContaContabil"
    sFrom = " FROM Produtos, EstoqueProduto, ProdutosFilial, Almoxarifado"
    sWhere = " WHERE ProdutosFilial.Produto = Produtos.Codigo AND EstoqueProduto.Almoxarifado = ProdutosFilial.Almoxarifado AND Almoxarifado.Codigo = ProdutosFilial.Almoxarifado AND EstoqueProduto.Produto = Produtos.Codigo AND Produtos.Compras = " & PRODUTO_COMPRAVEL & "AND Produtos.Ativo = " & PRODUTO_ATIVO & " AND ProdutosFilial.FilialEmpresa = ? " & IIf(objGeracaoPtoPedido.iExibePPZerado = MARCADO, "", "AND ProdutosFilial.PontoPedido > 0")
    sOrderBy = " ORDER BY Produtos.Codigo"
    
    If Len(Trim(objGeracaoPtoPedido.sCodProdutoDe)) > 0 Then
        
        sWhere = sWhere & " AND Produtos.Codigo >= ?"
    
    End If
    
    If Len(Trim(objGeracaoPtoPedido.sCodProdutoAte)) > 0 Then
        
        sWhere = sWhere & " AND Produtos.Codigo <= ?"
    
    End If
    
    For Each objTipoProduto In objGeracaoPtoPedido.colTipoProduto
        
        iIndice = iIndice + 1
            
        If iIndice = 1 Then
            sWhere = sWhere & " AND ("
        Else
            sWhere = sWhere & " OR "
        End If
        
        sWhere = sWhere & " Produtos.Tipo = " & objTipoProduto.iTipo
    
    Next
    
    If objGeracaoPtoPedido.colTipoProduto.Count > 0 Then
        sWhere = sWhere & " )"
    End If
    
    sSelect = "SELECT DISTINCT " & sFields & sFrom & sWhere & sOrderBy

End Sub

Function ProdutoFilial_Le_PontoPedido2(ByVal lComando As Long, ByVal sSelect As String, vCodigoP As Variant, vDescricaoP As Variant, vSiglaUMEstoqueP As Variant, vPontoPedidoPF As Variant, vEstoqueSegurancaPF As Variant, vConsumoMedioPF As Variant, vIntRessupPF As Variant, vLoteEconomicoPF As Variant, vAlmoxarifadoPF As Variant, vContaContabilEP As Variant, vFornecedorPF As Variant, vFilialFornPF As Variant, vFilialEmpresaPF As Variant, vContaContabilA As Variant) As Long
'Parte Fixa do SELECT

Dim lErro As Long
Dim ret As Integer

On Error GoTo Erro_ProdutoFilial_Le_PontoPedido2

    
    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 67422

    ret = Comando_BindVarInt(lComando, vCodigoP)
    If (ret <> AD_SQL_SUCESSO) Then gError 67423

    ret = Comando_BindVarInt(lComando, vDescricaoP)
    If (ret <> AD_SQL_SUCESSO) Then gError 67424

    ret = Comando_BindVarInt(lComando, vSiglaUMEstoqueP)
    If (ret <> AD_SQL_SUCESSO) Then gError 67425

    ret = Comando_BindVarInt(lComando, vPontoPedidoPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67426

    ret = Comando_BindVarInt(lComando, vEstoqueSegurancaPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67426

    ret = Comando_BindVarInt(lComando, vConsumoMedioPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67427

    ret = Comando_BindVarInt(lComando, vIntRessupPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67428

    ret = Comando_BindVarInt(lComando, vLoteEconomicoPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67429

    ret = Comando_BindVarInt(lComando, vAlmoxarifadoPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67430

    ret = Comando_BindVarInt(lComando, vContaContabilEP)
    If (ret <> AD_SQL_SUCESSO) Then gError 67431

    ret = Comando_BindVarInt(lComando, vFornecedorPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67432

    ret = Comando_BindVarInt(lComando, vFilialFornPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67433

    ret = Comando_BindVarInt(lComando, vFilialEmpresaPF)
    If (ret <> AD_SQL_SUCESSO) Then gError 67434
    
    ret = Comando_BindVarInt(lComando, vContaContabilA)
    If (ret <> AD_SQL_SUCESSO) Then gError 67435

    ProdutoFilial_Le_PontoPedido2 = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le_PontoPedido2:

    ProdutoFilial_Le_PontoPedido2 = gErr

    Select Case gErr

        Case 67422 To 67435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERACAOPTOPEDIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147678)

    End Select

    Exit Function

End Function

Function ProdutoFilial_Le_PontoPedido3(ByVal lComando As Long, ByVal objGeracaoPtoPedido As ClassGeracaoPtoPedido) As Long
'Parte Dinâmica do SELECT

Dim lErro As Long
Dim vFilialEmpresa As Variant
Dim vProdutoDe As Variant
Dim vProdutoAte As Variant
Dim vTipoProduto As Variant
Dim objTipoDeProduto As ClassTipoDeProduto

On Error GoTo Erro_ProdutoFilial_Le_PontoPedido3

    vFilialEmpresa = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, vFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67436

    If Len(Trim(objGeracaoPtoPedido.sCodProdutoDe)) > 0 Then
        vProdutoDe = objGeracaoPtoPedido.sCodProdutoDe
        lErro = Comando_BindVarInt(lComando, vProdutoDe)
        If lErro <> AD_SQL_SUCESSO Then gError 67437
    End If

    If Len(Trim(objGeracaoPtoPedido.sCodProdutoAte)) > 0 Then
        vProdutoAte = objGeracaoPtoPedido.sCodProdutoAte
        lErro = Comando_BindVarInt(lComando, vProdutoAte)
        If lErro <> AD_SQL_SUCESSO Then gError 67438
    End If

''    For Each objTipoDeProduto In objGeracaoPtoPedido.colTipoProduto
''
''        vTipoProduto = objTipoDeProduto.iTipo
''        lErro = Comando_BindVarInt(lComando, vTipoProduto)
''        If lErro <> AD_SQL_SUCESSO Then gError 67439
''
''    Next
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 67439

    ProdutoFilial_Le_PontoPedido3 = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le_PontoPedido3:

    ProdutoFilial_Le_PontoPedido3 = gErr

    Select Case gErr

        Case 67436 To 67439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147679)

    End Select

    Exit Function

End Function

Function ProdutoFilial_Le_PontoPedido4(ByVal lComando As Long, ByVal objGeracaoPtoPedido As ClassGeracaoPtoPedido, sCodigoP As String, sDescricaoP As String, sSiglaUMEstoqueP As String, dPontoPedidoPF As Double, dEstoqueSegurancaPF As Double, dConsumoMedioPF As Double, iIntRessupPF As Integer, dLoteEconomicoPF As Double, iAlmoxarifadoPF As Integer, sContaContabilEP As String, lFornecedorPF As Long, iFilialFornPF As Integer, iFilialEmpresaPF As Integer, sContaContabilA As String) As Long
'Executa o SELECT e Lê Registros

Dim lErro As Long
Dim objProdutoPtoPedido As ClassProdutoPtoPedido
Dim iIndice  As Integer

On Error GoTo Erro_ProdutoFilial_Le_PontoPedido4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67441

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67442

    Set objGeracaoPtoPedido.colProdutoPtoPedido = New Collection
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objProdutoPtoPedido = New ClassProdutoPtoPedido
        Set objProdutoPtoPedido.objProduto = New ClassProduto
        Set objProdutoPtoPedido.objEstoqueProduto = New ClassEstoqueProduto
        Set objProdutoPtoPedido.objProdutoFilial = New ClassProdutoFilial
        
        objProdutoPtoPedido.objProduto.sCodigo = sCodigoP
        objProdutoPtoPedido.objProduto.sDescricao = sDescricaoP
        objProdutoPtoPedido.objProduto.sSiglaUMEstoque = sSiglaUMEstoqueP
        
        With objProdutoPtoPedido.objProdutoFilial
            .dPontoPedido = dPontoPedidoPF
            .dEstoqueSeguranca = dEstoqueSegurancaPF
            .dConsumoMedio = dConsumoMedioPF
            .iIntRessup = iIntRessupPF
            .dLoteEconomico = dLoteEconomicoPF
            .iAlmoxarifado = iAlmoxarifadoPF
            .lFornecedor = lFornecedorPF
            .iFilialForn = iFilialFornPF
            .iFilialEmpresa = iFilialEmpresaPF
            .dTempoRessupMax = iIntRessupPF
        End With
        
        If Len(Trim(sContaContabilEP)) > 0 Then
            objProdutoPtoPedido.objEstoqueProduto.sContaContabil = sContaContabilEP
        Else
            objProdutoPtoPedido.objEstoqueProduto.sContaContabil = sContaContabilA
        End If
        
        'Preenche o Produto em Estoque Produto e Passa com parametro
        objProdutoPtoPedido.objEstoqueProduto.sProduto = objProdutoPtoPedido.objProduto.sCodigo
        
        'Le a soma das quantidades para o Produto - Filial em todos os Almoxarifados
        lErro = CF("EstoqueProduto_Le_Todos_AlmProdFilial", objProdutoPtoPedido.objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 67448 Then gError 67444
        
        objGeracaoPtoPedido.colProdutoPtoPedido.Add objProdutoPtoPedido

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67443

    Loop

    ProdutoFilial_Le_PontoPedido4 = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le_PontoPedido4:

    ProdutoFilial_Le_PontoPedido4 = gErr

    Select Case gErr

    Case 67441, 67443
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERACAOPTOPEDIDO", gErr)

    Case 67442, 67444

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147680)

    End Select

    Exit Function

End Function

Function EstoqueProduto_Le_Todos_AlmProdFilial(ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Lê a soma das Quantidades de todos Almoxarifados que pertence ao Produto passado em objEstoqueProduto

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueProduto As typeEstoqueProduto
Dim dOp As Double

On Error GoTo Erro_EstoqueProduto_Le_Todos_AlmProdFilial

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67445
    
    lErro = Comando_Executar(lComando, "SELECT SUM(QuantDispNossa), SUM(QuantReservada), MAX(DataInventario), SUM(QuantidadeInicial), SUM(SaldoInicial), SUM(QuantEmpenhada), SUM(QuantPedida), SUM(QuantRecIndl), SUM(QuantIndOutras), SUM(QuantDefeituosa), SUM(QuantConsig3), SUM(QuantConsig), SUM(QuantDemo3), SUM(QuantDemo), SUM(QuantConserto3), SUM(QuantConserto), SUM(QuantOutras3), SUM(QuantOutras), SUM(QuantBenef), SUM(QuantBenef3), SUM(QuantOP) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND  EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ?", _
        tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantPedido, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantBenef, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantOP, objEstoqueProduto.sProduto, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67446

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67447

    'Não encontrou ---> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67448
    
    objEstoqueProduto.dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
    objEstoqueProduto.dQuantReservada = tEstoqueProduto.dQuantReservada
    objEstoqueProduto.dtDataInventario = tEstoqueProduto.dtDataInventario
    objEstoqueProduto.dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
    objEstoqueProduto.dSaldoInicial = tEstoqueProduto.dSaldoInicial
    objEstoqueProduto.dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
    objEstoqueProduto.dQuantPedido = tEstoqueProduto.dQuantPedido
    objEstoqueProduto.dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
    objEstoqueProduto.dQuantInd = tEstoqueProduto.dQuantInd
    objEstoqueProduto.dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
    objEstoqueProduto.dQuantConsig3 = tEstoqueProduto.dQuantConsig3
    objEstoqueProduto.dQuantConsig = tEstoqueProduto.dQuantConsig
    objEstoqueProduto.dQuantDemo3 = tEstoqueProduto.dQuantDemo3
    objEstoqueProduto.dQuantDemo = tEstoqueProduto.dQuantDemo
    objEstoqueProduto.dQuantConserto3 = tEstoqueProduto.dQuantConserto3
    objEstoqueProduto.dQuantConserto = tEstoqueProduto.dQuantConserto
    objEstoqueProduto.dQuantOutras3 = tEstoqueProduto.dQuantOutras3
    objEstoqueProduto.dQuantOutras = tEstoqueProduto.dQuantOutras
    objEstoqueProduto.dQuantBenef = tEstoqueProduto.dQuantBenef
    objEstoqueProduto.dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    objEstoqueProduto.dQuantOP = tEstoqueProduto.dQuantOP
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueProduto_Le_Todos_AlmProdFilial = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Todos_AlmProdFilial:

    EstoqueProduto_Le_Todos_AlmProdFilial = gErr

    Select Case gErr
        
        Case 67445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67446, 67447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO3", gErr)

        Case 67448 'Não encontrou --> Tratar na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147681)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Motivo_Le_Codigo(ByVal objMotivo As AdmCodigoNome)
'Lê o Motivo com o código passado cadastrado na tabela Motivo

Dim lErro As Long
Dim lComando As Long
Dim sMotivo As String

On Error GoTo Erro_Motivo_Le_Codigo

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66785

    sMotivo = String(STRING_MOTIVO_ESCOLHA, 0)

    'Lê a Descrição do Motivo da tabela Motivo
    lErro = Comando_Executar(lComando, "SELECT Motivo FROM Motivo", sMotivo)
    If lErro <> AD_SQL_SUCESSO Then gError 66774

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66775

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66776

    objMotivo.sNome = sMotivo

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Motivo_Le_Codigo = SUCESSO

    Exit Function

Erro_Motivo_Le_Codigo:

    Motivo_Le_Codigo = gErr

    Select Case gErr

        Case 66785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66774, 66775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOTIVO", gErr)

        Case 66776 'Motivo não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147682)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Concorrencia_Le(ByVal objConcorrencia As ClassConcorrencia) As Long
'Lê dados da concorrência passada pelo código e filialempresa em objConcorrencia

Dim lErro As Long
Dim lComando As Long
Dim tConcorrencia As typeConcorrencia

On Error GoTo Erro_Concorrencia_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66785

    tConcorrencia.sDescricao = String(STRING_NOME, 0)

    'Lê dados da concorrência
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Descricao, Data, Comprador, TaxaFinanceira, TipoDestino, FornCliDestino, FilialDestino FROM Concorrencia WHERE Codigo  = ? AND FilialEmpresa = ?", _
    tConcorrencia.lNumIntDoc, tConcorrencia.iFilialEmpresa, tConcorrencia.sDescricao, tConcorrencia.dtData, tConcorrencia.iComprador, tConcorrencia.dTaxaFinanceira, tConcorrencia.iTipoDestino, tConcorrencia.lFornCliDestino, tConcorrencia.iFilialDestino, objConcorrencia.lCodigo, objConcorrencia.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66786

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66787

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66788

    objConcorrencia.lNumIntDoc = tConcorrencia.lNumIntDoc
    objConcorrencia.iFilialEmpresa = tConcorrencia.iFilialEmpresa
    objConcorrencia.sDescricao = tConcorrencia.sDescricao
    objConcorrencia.dtData = tConcorrencia.dtData
    objConcorrencia.iComprador = tConcorrencia.iComprador
    objConcorrencia.dTaxaFinanceira = tConcorrencia.dTaxaFinanceira
    objConcorrencia.iTipoDestino = tConcorrencia.iTipoDestino
    objConcorrencia.lFornCliDestino = tConcorrencia.lFornCliDestino
    objConcorrencia.iFilialDestino = tConcorrencia.iFilialDestino

    'Lê os itens da Concorrência
    lErro = ItensConcorrencia_Le(objConcorrencia)
    If lErro <> SUCESSO Then gError 51251

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Concorrencia_Le = SUCESSO

    Exit Function

Erro_Concorrencia_Le:

    Concorrencia_Le = gErr

    Select Case gErr

        Case 66785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66786, 66787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case 51251, 66788 'Concorrência não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147683)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Alcada_Avalia(ByVal sCodUsuario As String, ByVal dValor As Double, ByVal dtData As Date) As Long
'Verifica se o usuario possui alcada

Dim lErro As Long
Dim objAlcada As New ClassAlcada
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim dLiberado As Double

On Error GoTo Erro_Alcada_Avalia

    objAlcada.sCodUsuario = sCodUsuario

    lErro = CF("Alcada_Le", objAlcada)
    If lErro <> SUCESSO And lErro <> 49208 Then Error 49256
    ' O usuario nao possui alcada
    If lErro = 49208 Then Error 49257

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49284
    Next

    'Verifica se o valor ultrapassa o limite de operacao
    If objAlcada.dLimiteOperacao <= dValor Then Error 49258

    'Verifica se o valor ultrapassa o limite mensal
    If objAlcada.dLimiteMensal <= dValor Then Error 49259

    'Busca valor liberado do mes em questao
    lErro = Comando_Executar(alComando(0), "SELECT Liberado" & CStr(Month(dtData)) & " FROM ValorPCLiberado WHERE CodUsuario = ? And  Ano = ?", dLiberado, sCodUsuario, Year(dtData))
    If lErro <> AD_SQL_SUCESSO Then Error 49268

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49269

    If lErro = AD_SQL_SUCESSO Then

        'Verifica se o limite mensal comporta o que ja foi liberado mais o valor a liberar
        If objAlcada.dLimiteMensal < (dValor + dLiberado) Then Error 49270

    End If

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Alcada_Avalia = SUCESSO

    Exit Function

Erro_Alcada_Avalia:

    Alcada_Avalia = Err

    Select Case Err

        Case 49256

        Case 49257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_SEM_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITE_OPERACAO_MENOR", Err, objAlcada.dLimiteOperacao)

        Case 49259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITE_MENSAL_MENOR", Err, objAlcada.dLimiteMensal)

        Case 49268, 49269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALORPCLIBERADO", Err)

        Case 49270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITE_MENSAL_ULTRAPASSADO", Err, objAlcada.dLimiteMensal, objAlcada.sCodUsuario)

        Case 49284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147684)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ComprasConfig_Le_Conteudo(ByVal objComprasConfig As ClassComprasConfig) As Long
 'Lê o registro em ComprasConfig de acordo com o Código passado.

Dim lErro As Long
Dim sConteudo As String
Dim lComando As Long

On Error GoTo Erro_ComprasConfig_Le_Conteudo

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 64259

    sConteudo = String(STRING_CONTEUDO, 0)

    'Ler registo
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM ComprasConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 64260

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64261

    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 64262

    objComprasConfig.sConteudo = sConteudo

    Call Comando_Fechar(lComando)
    
    ComprasConfig_Le_Conteudo = SUCESSO

    Exit Function

Erro_ComprasConfig_Le_Conteudo:

    ComprasConfig_Le_Conteudo = Err

    Select Case Err
        
        Case 64259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64260, 64261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", Err, objComprasConfig.sCodigo)

        Case 64262
            Call Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO", Err, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147685)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemPedCotacao_Le_PedidoCotacao(ByVal objItemPedCotacao As ClassItemPedCotacao, ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê o Código do Pedido de Cotacao a que pertence o Item de Pedido de Cotacao passado como parametro

Dim lErro As Long
Dim lTransacao As Long
Dim alComando As Long
Dim lPedidoCotacao As Long
Dim tPedCotacao As typePedidoCotacao

On Error GoTo Erro_ItemPedCotacao_Le_PedidoCotacao

    'Abre o comando
    alComando = Comando_Abrir
    If alComando = 0 Then gError 68494

    'Lê ItemPedCotacao a partir do NumIntDoc passado como parametro
    lErro = Comando_Executar(alComando, "SELECT CotacaoProduto,PedidoCotacao FROM ItemPedCotacao WHERE NumIntDoc=?", objItemPedCotacao.lCotacaoProduto, lPedidoCotacao, objItemPedCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 68487
    
    lErro = Comando_BuscarPrimeiro(alComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68488
    
    If lErro = AD_SQL_SEM_DADOS Then gError 68489
    
    'Busca no BD o Pedido de Cotacao a partir do NumIntDoc
    lErro = Comando_Executar(alComando, "SELECT NumIntDoc,FilialEmpresa,Codigo FROM PedidoCotacao WHERE NumIntDoc=?", tPedCotacao.lNumIntDoc, tPedCotacao.iFilialEmpresa, tPedCotacao.lCodigo, lPedidoCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 68490

    lErro = Comando_BuscarPrimeiro(alComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68491
    
    If lErro = AD_SQL_SEM_DADOS Then gError 68492
    
    objPedidoCotacao.lCodigo = tPedCotacao.lCodigo
    objPedidoCotacao.iFilialEmpresa = tPedCotacao.iFilialEmpresa
    
    Call Comando_Fechar(alComando)

    ItemPedCotacao_Le_PedidoCotacao = SUCESSO
    
    Exit Function
    
Erro_ItemPedCotacao_Le_PedidoCotacao:
    
    ItemPedCotacao_Le_PedidoCotacao = gErr
    
    Select Case gErr
    
        Case 68487, 68488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO1", gErr)
            
        Case 68489, 68492
            'Erros tratados nas rotinas chamadas
            
        Case 68490, 68491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO1", gErr)
            
        Case 68493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 68494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 68495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147686)
            
    End Select
    
    Call Comando_Fechar(alComando)

    lTransacao = Transacao_Rollback

    Exit Function
    
End Function

Function ItemNFItemPC_Le(ByVal lNumItemPC As Long, ByVal colItemNF As Collection) As Long
'Lê todos os itens de Nota Fiscal relacionados ao item de Pedido de compras passado

Dim lErro As Long
Dim lComando As Long
Dim objItemNF As ClassItemNF
Dim tItemNF As typeItemNF

On Error GoTo Erro_ItemNFItemPC_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66707

    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    tItemNF.sSerieNF = String(STRING_SERIE, 0)
    tItemNF.sProduto = String(STRING_PRODUTO, 0)

    'Pesquisa os Itens de Nota Fiscal relacionados ao Item de Pedido de Compras
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ItensNFiscal.Produto, ItensNFiscal.UnidadeMed, ItemNFItemPC.Quantidade, NFiscal.NumNotaFiscal, NFiscal.Serie, ItensNFiscal.Item FROM NFiscal, ItensNFiscal, ItemNFItemPC WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItemNFItemPC.ItemPedCompra = ?", _
        tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.lNumNotaFiscal, tItemNF.sSerieNF, tItemNF.iItem, lNumItemPC)
    If lErro <> AD_SQL_SUCESSO Then gError 66708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66709

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66711

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemNF = New ClassItemNF

        objItemNF.sUnidadeMed = tItemNF.sUnidadeMed
        objItemNF.dQuantidade = tItemNF.dQuantidade
        objItemNF.lNumNFOrig = tItemNF.lNumNotaFiscal
        objItemNF.sSerieNFOrig = tItemNF.sSerieNF
        objItemNF.sProduto = tItemNF.sProduto
        objItemNF.iItem = tItemNF.iItem

        'Adiciona na coleção
        colItemNF.Add objItemNF

        'Busca no BD o próximo item NF
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66710

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemNFItemPC_Le = SUCESSO

    Exit Function

Erro_ItemNFItemPC_Le:

    ItemNFItemPC_Le = gErr

    Select Case gErr

        Case 66707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66708, 66709, 66710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66711 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147687)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Le_Numero(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Busca no BD Pedido de Compra com FilialEmpresa e Codigo passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim lNumIntDoc As Long
Dim tPedCompra As typePedidoCompras

On Error GoTo Erro_PedidoCompra_Le_Numero

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56139

    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)
    tPedCompra.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)

    'OK  por que não leu o número interno da observação
    'Busca na tabela de Pedido de Compra o pedido de Compra com codigo e filial empresa passados
'leo
    lErro = Comando_Executar(lComando, "SELECT Embalagem,Taxa,Moeda,NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ValorProdutos,ObsEmbalagem, Observacao, DataRefFluxo, DataBaixa, TabelaPreco, UsuReg, UsuRegAprov, UsuRegEnvio, DataRegAprov FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa=?", _
        tPedCompra.iEmbalagem, tPedCompra.dTaxa, tPedCompra.iMoeda, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.dValorProdutos, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dtDataRefFluxo, tPedCompra.dtDataBaixa, tPedCompra.iTabelaPreco, tPedCompra.sUsuReg, tPedCompra.sUsuRegAprov, tPedCompra.sUsuRegEnvio, tPedCompra.dtDataRegAprov, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 56140

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56141

    'Se nao encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56142

    'Preenche objPedidoCompra com o que foi lido do banco de dados
    objPedidoCompra.iMoeda = tPedCompra.iMoeda
    objPedidoCompra.iEmbalagem = tPedCompra.iEmbalagem
    objPedidoCompra.dTaxa = tPedCompra.dTaxa
    objPedidoCompra.lNumIntDoc = tPedCompra.lNumIntDoc
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataBaixa = DATA_NULA
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra.lObservacao = tPedCompra.lObservacao  'OK  Por que não leu o número interno da observação
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.dValorProdutos = tPedCompra.dValorProdutos
    objPedidoCompra.dtDataRefFluxo = tPedCompra.dtDataRefFluxo
    objPedidoCompra.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra.iTabelaPreco = tPedCompra.iTabelaPreco
    objPedidoCompra.sUsuReg = tPedCompra.sUsuReg
    objPedidoCompra.sUsuRegAprov = tPedCompra.sUsuRegAprov
    objPedidoCompra.sUsuRegEnvio = tPedCompra.sUsuRegEnvio
    objPedidoCompra.dtDataRegAprov = tPedCompra.dtDataRegAprov

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedidoCompra_Le_Numero = SUCESSO

    Exit Function

Erro_PedidoCompra_Le_Numero:

    PedidoCompra_Le_Numero = Err

    Select Case Err

        Case 56139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56140, 56141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra.lCodigo)

        Case 56142

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147688)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Le_Todos(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Busca no BD Pedido de Compra com FilialEmpresa e Codigo passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim lNumIntDoc As Long
Dim tPedCompra As typePedidoCompras
Dim iStatus As Integer

On Error GoTo Erro_PedidoCompra_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 68483

    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)
    tPedCompra.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tPedCompra.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)

    'OK  por que não leu o número interno da observação
    'Busca na tabela de Pedido de Compra o pedido de Compra com codigo e filial empresa passados
    If objPedidoCompra.lCodigo <> 0 Then
        If objPedidoCompra.iFilialEmpresa = EMPRESA_TODA Then objPedidoCompra.iFilialEmpresa = giFilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ValorProdutos,ObsEmbalagem, Observacao, DataBaixa, Status, UsuReg, UsuRegAprov, UsuRegEnvio, DataRegAprov FROM PedidoCompraTodos WHERE Codigo = ? AND FilialEmpresa=?", _
        tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.dValorProdutos, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dtDataBaixa, iStatus, tPedCompra.sUsuReg, tPedCompra.sUsuRegAprov, tPedCompra.sUsuRegEnvio, tPedCompra.dtDataRegAprov, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ValorProdutos,ObsEmbalagem, Observacao, DataBaixa, Status, UsuReg, UsuRegAprov, UsuRegEnvio, DataRegAprov FROM PedidoCompraTodos WHERE NumIntdoc=?", _
        tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.dValorProdutos, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dtDataBaixa, iStatus, tPedCompra.sUsuReg, tPedCompra.sUsuRegAprov, tPedCompra.sUsuRegEnvio, tPedCompra.dtDataRegAprov, objPedidoCompra.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 68484

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68485

    'Se nao encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 68486

    'Preenche objPedidoCompra com o que foi lido do banco de dados
    objPedidoCompra.lNumIntDoc = tPedCompra.lNumIntDoc
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra.lObservacao = tPedCompra.lObservacao
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.dValorProdutos = tPedCompra.dValorProdutos
    objPedidoCompra.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra.iStatus = iStatus
    objPedidoCompra.sUsuReg = tPedCompra.sUsuReg
    objPedidoCompra.sUsuRegAprov = tPedCompra.sUsuRegAprov
    objPedidoCompra.sUsuRegEnvio = tPedCompra.sUsuRegEnvio
    objPedidoCompra.dtDataRegAprov = tPedCompra.dtDataRegAprov
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedidoCompra_Le_Todos = SUCESSO

    Exit Function

Erro_PedidoCompra_Le_Todos:

    PedidoCompra_Le_Todos = gErr

    Select Case gErr

        Case 68483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68484, 68485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 68486
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147689)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItensPC_LeTodos(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê no BD a coleção de ítens do Pedido de Compra e devolve em colItensPedido dentro de objPedidoCompra

Dim lErro As Long
Dim tItemPedido As typeItemPedCompra
Dim lComando As Long
Dim lTransacao As Long
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_ItensPC_LeTodos

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 68431

    tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0) 'OK o tamanho correto da string é 255
    tItemPedido.sUM = String(STRING_UM_SIGLA, 0)

    'Le da tabela de ItensPedCompraTodos a colecao de itens do Pedido de Compra passado como parametro
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS,DeliveryDate, TempoTransito FROM ItensPedCompraTodos WHERE PedCompra = ? ORDER BY NumIntDoc", tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, _
    tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, tItemPedido.dtDeliveryDate, tItemPedido.iTempoTransito, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 68432

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68433

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 68434

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPC = New ClassItemPedCompra

        'preenchendo objitem
        objItemPC.dAliquotaICMS = tItemPedido.dAliquotaICMS
        objItemPC.dAliquotaIPI = tItemPedido.dAliquotaIPI
        objItemPC.dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        objItemPC.dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        objItemPC.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemPC.dQuantidade = tItemPedido.dQuantidade
        objItemPC.dQuantRecebida = tItemPedido.dQuantRecebida
        objItemPC.dQuantRecebimento = tItemPedido.dQuantRecebimento
        objItemPC.dtDataLimite = tItemPedido.dtDataLimite
        objItemPC.dValorDesconto = tItemPedido.dValorDesconto
        objItemPC.dValorIPI = tItemPedido.dValorIPI
        objItemPC.iRebebForaFaixa = tItemPedido.iRebebForaFaixa
        objItemPC.iStatus = tItemPedido.iStatus
        objItemPC.iTipoOrigem = tItemPedido.iTipoOrigem
        objItemPC.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemPC.lNumIntOrigem = tItemPedido.lNumIntOrigem
        objItemPC.lObservacao = tItemPedido.lObservacao
        objItemPC.sDescProduto = tItemPedido.sDescProduto
        objItemPC.sProduto = tItemPedido.sProduto
        objItemPC.sUM = tItemPedido.sUM

        objItemPC.dtDeliveryDate = tItemPedido.dtDeliveryDate
        objItemPC.iTempoTransito = tItemPedido.iTempoTransito
        
        lErro = CF("LocalizacaoItemPC_Le", objItemPC)
        If lErro <> SUCESSO And lErro <> 56361 Then gError 68435

        'Insere em colItens os valores de objItemPC
        objPedidoCompra.colItens.Add objItemPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68436
    Loop

    '######################################################
    'Inserido por Wagner 17/07/2006
    lErro = CF("PCItens_Le_Customizado", objPedidoCompra)
    If lErro <> SUCESSO Then gError 181225
    '######################################################

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPC_LeTodos = SUCESSO

    Exit Function

Erro_ItensPC_LeTodos:

    ItensPC_LeTodos = gErr

    Select Case gErr

        Case 68431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68432, 68433, 68436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 68435, 181225
            'Erro tratado na rotina chamadora

        Case 68434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objPedidoCompra.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147690)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompraBaixado_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra,cujo numero interno foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim tPedCompra As typePedidoCompras

On Error GoTo Erro_PedidoCompraBaixado_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89234
    
    'OK  Isso aqui não compila. Essa variável não foi declarada
    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)

    'Busca na tabela de Pedido de Compra o pedido de Compra com codigo e filial empresa passados
    lErro = Comando_Executar(lComando, "SELECT Embalagem,Taxa,Moeda,NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao, DataRefFluxo, DataBaixa, TabelaPreco FROM PedidoCompraBaixado WHERE NumIntDoc = ?", _
        tPedCompra.iEmbalagem, tPedCompra.dTaxa, tPedCompra.iMoeda, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dtDataRefFluxo, tPedCompra.dtDataBaixa, tPedCompra.iTabelaPreco, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89235

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89236

    'PedidoCompra nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 89237

    'carrega os dados em objPedidoCompra
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.lObservacao = tPedCompra.lObservacao
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.iEmbalagem = tPedCompra.iEmbalagem
    objPedidoCompra.dTaxa = tPedCompra.dTaxa
    objPedidoCompra.iMoeda = tPedCompra.iMoeda
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra.dtDataRefFluxo = tPedCompra.dtDataRefFluxo
    objPedidoCompra.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra.iTabelaPreco = tPedCompra.iTabelaPreco
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedidoCompraBaixado_Le = SUCESSO

    Exit Function

Erro_PedidoCompraBaixado_Le:

    PedidoCompraBaixado_Le = gErr

    Select Case gErr

        Case 89234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89235, 89236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 89237
            'erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147691)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompraBaixado_Le_Numero(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Busca no BD Pedido de Compra Baixado para a FilialEmpresa e Codigo passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim tPedCompra As typePedidoCompras

On Error GoTo Erro_PedidoCompraBaixado_Le_Numero

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56134

    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)

    'Busca no BD Pedido de Compra Baixado com Codigo e FilialEmpresa passados como parametros
    lErro = Comando_Executar(lComando, "SELECT Embalagem,Taxa,Moeda,NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao, DataRefFluxo, DataBaixa, TabelaPreco FROM PedidoCompraBaixado WHERE Codigo = ? AND FilialEmpresa=?", tPedCompra.iEmbalagem, tPedCompra.dTaxa, tPedCompra.iMoeda, _
                tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, tPedCompra.sTipoFrete, tPedCompra.iTransportadora, _
                tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.dtDataRefFluxo, tPedCompra.dtDataBaixa, tPedCompra.iTabelaPreco, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 56135

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56136

    'Se encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56137

    'Preenche objPedidoCompra com o que foi lido do banco de dados
    objPedidoCompra.lNumIntDoc = tPedCompra.lNumIntDoc
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.lObservacao = tPedCompra.lObservacao  'OK  Por que não leu o número interno da observação
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.dValorProdutos = tPedCompra.dValorProdutos
    objPedidoCompra.iEmbalagem = tPedCompra.iEmbalagem
    objPedidoCompra.dTaxa = tPedCompra.dTaxa
    objPedidoCompra.iMoeda = tPedCompra.iMoeda
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra.dtDataRefFluxo = tPedCompra.dtDataRefFluxo
    objPedidoCompra.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra.iTabelaPreco = tPedCompra.iTabelaPreco

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedidoCompraBaixado_Le_Numero = SUCESSO

    Exit Function

Erro_PedidoCompraBaixado_Le_Numero:

    PedidoCompraBaixado_Le_Numero = Err

    Select Case Err

        Case 56134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56135, 56136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA_BAIXADO", Err, objPedidoCompra.lCodigo)

        Case 56137

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147692)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCompra_Lock(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Le o numero interno dos itens do pedido, guarda na colecao de pedidos e faz lock exclusive dos itens

Dim lErro As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim objItemPC As ClassItemPedCompra
Dim lComando As Long
Dim lNumIntOrigem As Long

On Error GoTo Erro_ItensPedCompra_Lock

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56329

    'Busca na tabela de ItensPedCompra os itens do Pedido de Compra passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,PedCompra,NumIntOrigem FROM ItensPedCompra WHERE  PedCompra = ?", 0, lNumIntDoc, lCodigo, lNumIntOrigem, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 56330

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56331

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56332

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPC = New ClassItemPedCompra
        objItemPC.lNumIntDoc = lNumIntDoc

        objItemPC.lNumIntOrigem = lNumIntOrigem

        'Lock Exclusive nos itens do Pedido de Compra
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 56333

        'Adiciona os itens no colecao de itens
        objPedidoCompra.colItens.Add objItemPC

        'Busca o proximo item do Pedido de Compra passado como parametro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56371

    Loop

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    ItensPedCompra_Lock = SUCESSO

    Exit Function

Erro_ItensPedCompra_Lock:

    ItensPedCompra_Lock = Err

    Select Case Err

        Case 56329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56330, 56331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", lErro, objPedidoCompra.lNumIntDoc)

        Case 56332, 56371

        Case 56333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDCOMPRA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147693)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BloqueiosPC_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Le o BloqueioPC do Pedido de Compra passado como parametro

Dim lErro As Long
Dim tBloqueioPC As typeBloqueioPC
Dim lComando As Long
Dim objBloqueioPC As New ClassBloqueioPC

On Error GoTo Erro_BloqueiosPC_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56350

    tBloqueioPC.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tBloqueioPC.sResponsavel = String(STRING_RESPONSAVEL_RESERVA, 0)
    tBloqueioPC.sCodUsuarioLib = String(STRING_USUARIO_CODIGO, 0)

    'Busca na tabela de BloqueioPC, os bloqueios  do Pedido de Compra passado como parametro
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa,PedCompras,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,DataLib FROM BloqueiosPC WHERE FilialEmpresa = ? AND PedCompras = ? ", tBloqueioPC.iFilialEmpresa, tBloqueioPC.lPedCompras, tBloqueioPC.iSequencial, tBloqueioPC.iTipoDeBloqueio, tBloqueioPC.sCodUsuario, tBloqueioPC.sResponsavel, _
    tBloqueioPC.dtData, tBloqueioPC.sCodUsuarioLib, tBloqueioPC.dtDataLib, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 56352

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56353


    Do While lErro = AD_SQL_SUCESSO

        Set objBloqueioPC = New ClassBloqueioPC

        'Preenche objBloqueioPC
        objBloqueioPC.dtDataLib = tBloqueioPC.dtDataLib
        objBloqueioPC.iFilialEmpresa = tBloqueioPC.iFilialEmpresa
        objBloqueioPC.dtData = tBloqueioPC.dtData
        objBloqueioPC.iSequencial = tBloqueioPC.iSequencial
        objBloqueioPC.iTipoBloqueio = tBloqueioPC.iTipoDeBloqueio
        objBloqueioPC.sCodUsuario = tBloqueioPC.sCodUsuario
        objBloqueioPC.sCodUsuarioLib = tBloqueioPC.sCodUsuarioLib
        objBloqueioPC.sResponsavel = tBloqueioPC.sResponsavel
        objBloqueioPC.lPedCompras = tBloqueioPC.lPedCompras

        'Insere em colBloqueiosPC o conteudo de objBloqueioPC
        objPedidoCompra.colBloqueiosPC.Add objBloqueioPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56355
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    BloqueiosPC_Le = SUCESSO

    Exit Function

Erro_BloqueiosPC_Le:

    BloqueiosPC_Le = Err

    Select Case Err

        Case 56350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56352, 56353, 56355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC", Err, objPedidoCompra.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147694)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCotacaoTodos_Le(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê os ítens do Pedido de Cotação baixado ou não

Dim iIndice As Integer
Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItensPedCotacao As typeItemPedCotacao
Dim colAlocacoes As ColAlocacoesItemNF
Dim tItensCotacao As typeItemCotacao
Dim tCotacaoProduto As typePedidoCotacao
Dim objItemPedidoCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao

On Error GoTo Erro_ItensPedCotacaoTodos_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67750

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 67751

    Set objPedidoCotacao.colItens = New Collection
    
    tItensPedCotacao.sProduto = String(STRING_PRODUTO, 0)
    tItensPedCotacao.sUM = String(STRING_UM_SIGLA, 0)
    tItensPedCotacao.sObservacao = String(STRING_ITEMPEDCOTACAO_OBSERVACAO, 0)

    'Pesquisa os ítens de PedidoCotacao
    lErro = Comando_Executar(lComando, "SELECT ItemPedCotacaoTodos.Observacao, ItemPedCotacaoTodos.NumIntDoc, ItemPedCotacaoTodos.CotacaoProduto, CotacaoProdutoTodas.NumIntDoc, CotacaoProdutoTodas.Produto, CotacaoProdutoTodas.Quantidade, CotacaoProdutoTodas.UM, CotacaoProdutoTodas.Fornecedor, CotacaoProdutoTodas.Filial, CotacaoProdutoTodas.Cotacao FROM PedidoCotacaoTodos, ItemPedCotacaoTodos, CotacaoProdutoTodas WHERE PedidoCotacaoTodos.NumIntDoc = ItemPedCotacaoTodos.PedidoCotacao AND ItemPedCotacaoTodos.CotacaoProduto = CotacaoProdutoTodas.NumIntDoc AND PedidoCotacaoTodos.Codigo = ? AND PedidoCotacaoTodos.FilialEmpresa = ?", _
        tItensPedCotacao.sObservacao, tItensPedCotacao.lNumIntDoc, tItensPedCotacao.lCotacaoProduto, tCotacaoProduto.lNumIntDoc, tItensPedCotacao.sProduto, tItensPedCotacao.dQuantidade, tItensPedCotacao.sUM, tCotacaoProduto.lFornecedor, tCotacaoProduto.iFilial, tItensPedCotacao.lCotacaoProduto, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67752
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67753

    'Enquanto encontrar ItensPedCotacao
    Do While lErro <> AD_SQL_SEM_DADOS

        'Guarda na coleção de ítens de objPedidoCotacao
        Set objItemPedidoCotacao = New ClassItemPedCotacao

        objItemPedidoCotacao.lNumIntDoc = tItensPedCotacao.lNumIntDoc
        objPedidoCotacao.lCotacao = tItensPedCotacao.lCotacaoProduto
        objItemPedidoCotacao.sProduto = tItensPedCotacao.sProduto
        objItemPedidoCotacao.dQuantidade = tItensPedCotacao.dQuantidade
        objItemPedidoCotacao.sUM = tItensPedCotacao.sUM
        objItemPedidoCotacao.sObservacao = tItensPedCotacao.sObservacao
        
        'Lê ItensCotação (baixados ou não) vinculados ao ItemPedCotacao
        lErro = Comando_Executar(lComando1, "SELECT CondPagto, PrecoUnitario, OutrasDespesas, ValorFrete, ValorSeguro, ValorDesconto, ValorTotal, ValorIPI, AliquotaIPI, AliquotaICMS, QuantEntrega, Observacao, PrazoEntrega, Moeda, Taxa FROM ItensCotacaoTodos WHERE ItensCotacaoTodos.ItemPedCotacao = ?", tItensCotacao.iCondPagto, tItensCotacao.dPrecoUnitario, tItensCotacao.dOutrasDespesas, tItensCotacao.dValorFrete, tItensCotacao.dValorSeguro, tItensCotacao.dValorDesconto, tItensCotacao.dValorTotal, tItensCotacao.dValorIPI, tItensCotacao.dAliquotaIPI, tItensCotacao.dAliquotaICMS, tItensCotacao.dQuantEntrega, tItensCotacao.lObservacao, tItensCotacao.iPrazoEntrega, tItensCotacao.iMoeda, tItensCotacao.dTaxa, objItemPedidoCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 67755

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67756

        'Enquanto encontrar ItensCotacao
        Do While lErro <> AD_SQL_SEM_DADOS

            'Guarda na coleção de itens cotação de objItemPedidoCotacao
            Set objItensCotacao = New ClassItemCotacao

            objItensCotacao.iCondPagto = tItensCotacao.iCondPagto
            objItensCotacao.dPrecoUnitario = tItensCotacao.dPrecoUnitario
            objItensCotacao.dOutrasDespesas = tItensCotacao.dOutrasDespesas
            objItensCotacao.dValorFrete = tItensCotacao.dValorFrete
            objItensCotacao.dValorSeguro = tItensCotacao.dValorSeguro
            objItensCotacao.dValorDesconto = tItensCotacao.dValorDesconto
            objItensCotacao.dValorTotal = tItensCotacao.dValorTotal
            objItensCotacao.dValorIPI = tItensCotacao.dValorIPI
            objItensCotacao.dAliquotaIPI = tItensCotacao.dAliquotaIPI
            objItensCotacao.dAliquotaICMS = tItensCotacao.dAliquotaICMS
            objItensCotacao.dQuantEntrega = tItensCotacao.dQuantEntrega
            objItensCotacao.lObservacao = tItensCotacao.lObservacao
            objItensCotacao.iPrazoEntrega = tItensCotacao.iPrazoEntrega
            objItensCotacao.iMoeda = tItensCotacao.iMoeda
            objItensCotacao.dTaxa = tItensCotacao.dTaxa

            'Adiciona ItemCotação na coleção
            objItemPedidoCotacao.colItensCotacao.Add objItensCotacao

            'Busca próximo ItemCotação
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67754

        Loop

        'Adiciona ItemPedCotacao na coleção
        objPedidoCotacao.colItens.Add objItemPedidoCotacao

        'Busca próximo ItemPedCotação
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 67757

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPedCotacaoTodos_Le = SUCESSO

    Exit Function

Erro_ItensPedCotacaoTodos_Le:

    ItensPedCotacaoTodos_Le = gErr

    Select Case gErr

        Case 67750, 67751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67752, 67753, 67754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAOTODOS", gErr, objPedidoCotacao.lCodigo)

        Case 67755, 67756, 67757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAOTODOS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147695)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function CotacaoTodas_Le(ByVal objCotacao As ClassCotacao) As Long
'Lê a tabela Cotação (baixada ou não) passada pelo NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tCotacao As typeCotacao

On Error GoTo Erro_CotacaoTodas_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67758

    With tCotacao

        tCotacao.sDescricao = String(STRING_PEDIDO_COTACAO_DESCRICAO, 0)

        'Lê a cotação baixada ou não a patir do seu NumIntDoc
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, Descricao, Data, TipoDestino, FornCliDestino, FilialDestino, Comprador FROM CotacaoTodas WHERE NumIntDoc = ? ", tCotacao.iFilialEmpresa, tCotacao.lCodigo, tCotacao.sDescricao, tCotacao.dtData, tCotacao.iTipoDestino, tCotacao.lFornCliDestino, tCotacao.iFilialDestino, tCotacao.iComprador, objCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 67759

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67760

        'Se não encontrou a Cotação, Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67761

        'Armazena as informações lidas do bd no objeto
        objCotacao.iFilialEmpresa = tCotacao.iFilialEmpresa
        objCotacao.lCodigo = tCotacao.lCodigo
        objCotacao.sDescricao = tCotacao.sDescricao
        objCotacao.dtData = tCotacao.dtData
        objCotacao.iTipoDestino = tCotacao.iTipoDestino
        objCotacao.lFornCliDestino = tCotacao.lFornCliDestino
        objCotacao.iFilialDestino = tCotacao.iFilialDestino
        objCotacao.iComprador = tCotacao.iComprador

    End With

    'Fecha comandos
    Call Comando_Fechar(lComando)

    CotacaoTodas_Le = SUCESSO

    Exit Function

Erro_CotacaoTodas_Le:

    CotacaoTodas_Le = gErr

    Select Case gErr

        Case 67758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67759, 67760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOTODAS", gErr)

        Case 67761 'Não existe na tabela Cotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147696)

    End Select

    'Fecha comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCotacaoTodos_Le(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê o Pedido de Cotação (Baixado ou Não) a partir do Código e FilialEmpresa passados

Dim lErro As Long
Dim lComando As Long
Dim tPedidoCotacao As typePedidoCotacao

On Error GoTo Erro_PedidoCotacaoTodos_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67780

    With tPedidoCotacao

        .sContato = String(STRING_CONTATO, 0)

        'Lê os PedidoCotacao onde o Codigo e a FilialEmpresa forem iguais ao do objPedidoCotacao
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Fornecedor, Filial, Contato, DataEmissao, Data, DataValidade, TipoFrete,CondPagtoPrazo, DataBaixa FROM PedidoCotacaoTodos WHERE Codigo = ? AND FilialEmpresa=?", .lNumIntDoc, .lFornecedor, .iFilial, .sContato, .dtDataEmissao, .dtData, .dtDataValidade, .iTipoFrete, .iCondPagtoPrazo, .dtDataBaixa, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 67781

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67782
        
        'Se não encontrou o Pedido de Cotação, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67783

        'Armazena as informações lidas do bd no objeto
        objPedidoCotacao.lNumIntDoc = .lNumIntDoc
        objPedidoCotacao.lFornecedor = .lFornecedor
        objPedidoCotacao.iFilial = .iFilial
        objPedidoCotacao.sContato = .sContato
        objPedidoCotacao.dtDataEmissao = .dtDataEmissao
        objPedidoCotacao.dtData = .dtData
        objPedidoCotacao.dtDataValidade = .dtDataValidade
        objPedidoCotacao.iTipoFrete = .iTipoFrete
        objPedidoCotacao.iCondPagtoPrazo = .iCondPagtoPrazo
        objPedidoCotacao.dtDataBaixa = .dtDataBaixa

    End With

    'Fechamento dos Comandos
    Call Comando_Fechar(lComando)

    PedidoCotacaoTodos_Le = SUCESSO

    Exit Function

Erro_PedidoCotacaoTodos_Le:

    PedidoCotacaoTodos_Le = gErr

    Select Case gErr

        Case 67780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67781, 67782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAOTODOS", gErr)

        Case 67783 'Não existe na tabela de PedidoCotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147697)

    End Select

    'Fechamento dos Comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoProdutoItemRC_Le(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lCotacaoProduto As Long
Dim lItemReqCompra As Long
Dim dQuantidadeCotar As Double

On Error GoTo Erro_CotacaoProdutoItemRC_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 53784

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 53785

    'Verifica se o Pedido de Cotacao é avulso (não ligado a Requisição de Compra)
    lErro = Comando_Executar(lComando, "SELECT  CotacaoProdutoItemRC.CotacaoProduto, CotacaoProdutoItemRC.ItemReqCompra, CotacaoProdutoItemRC.QuantidadeCotar  FROM PedidoCotacao, ItemPedCotacao, CotacaoProduto, CotacaoProdutoItemRC WHERE PedidoCotacao.NumIntDoc=ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND CotacaoProduto.NumIntDoc=CotacaoProdutoItemRC.CotacaoProduto AND PedidoCotacao.Codigo = ? AND PedidoCotacao.FilialEmpresa=? ", lCotacaoProduto, lItemReqCompra, dQuantidadeCotar, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 53786

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53787
    'Se não encontrou --> erro.
    If lErro = AD_SQL_SEM_DADOS Then Error 53788

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    CotacaoProdutoItemRC_Le = SUCESSO

    Exit Function

Erro_CotacaoProdutoItemRC_Le:

    CotacaoProdutoItemRC_Le = Err

    Select Case Err

        Case 53784, 53785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53786, 53787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", Err)

        Case 53788 'Erro tratado na rotina que chamou a função.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147698)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function RequisicaoComprasTodas_Le(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê as Requisições de compras (baixadas ou não) associadas ao Pedidos de Compra passado

Dim lErro As Long
Dim lComando As Long
Dim tRequisicaoCompras As typeRequisicaoCompras

On Error GoTo Erro_RequisicaoComprasTodas_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66693

    tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
    tRequisicaoCompras.sDigitador = String(STRING_CODUSUARIO, 0)
    tRequisicaoCompras.sOPCodigo = String(STRING_OPCODIGO, 0)
    tRequisicaoCompras.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
        
    'Pesquisa as Requisições de Compras NumIntDoc passado
'    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo FROM RequisicaoCompraTodas WHERE RequisicaoCompraTodas.NumIntDoc = ? ", _
'        tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataEnvio, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sDigitador, tRequisicaoCompras.sCcl, tRequisicaoCompras.iFilialCompra, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, tRequisicaoCompras.iTipoTributacao, tRequisicaoCompras.sOPCodigo, objRequisicaoCompras.lNumIntDoc)
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, NumIntDocItemOP, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompraTodas WHERE RequisicaoCompraTodas.NumIntDoc = ? ", _
    tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataEnvio, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sDigitador, tRequisicaoCompras.sCcl, tRequisicaoCompras.iFilialCompra, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, tRequisicaoCompras.iTipoTributacao, tRequisicaoCompras.sOPCodigo, tRequisicaoCompras.lNumIntDocItemOP, tRequisicaoCompras.sUsuReg, tRequisicaoCompras.sUsuRegAprov, tRequisicaoCompras.sUsuRegBaixa, tRequisicaoCompras.sUsuRegEnvio, objRequisicaoCompras.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66694

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66695

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66696

    'Preenche o objRequisiçãoCompras
    objRequisicaoCompras.iFilialEmpresa = tRequisicaoCompras.iFilialEmpresa
    objRequisicaoCompras.lCodigo = tRequisicaoCompras.lCodigo
    objRequisicaoCompras.dtData = tRequisicaoCompras.dtData
    objRequisicaoCompras.dtDataEnvio = tRequisicaoCompras.dtDataEnvio
    objRequisicaoCompras.dtDataLimite = tRequisicaoCompras.dtDataLimite
    objRequisicaoCompras.lUrgente = tRequisicaoCompras.lUrgente
    objRequisicaoCompras.lRequisitante = tRequisicaoCompras.lRequisitante
    objRequisicaoCompras.sDigitador = tRequisicaoCompras.sDigitador
    objRequisicaoCompras.sCcl = tRequisicaoCompras.sCcl
    objRequisicaoCompras.iFilialCompra = tRequisicaoCompras.iFilialCompra
    objRequisicaoCompras.iTipoDestino = tRequisicaoCompras.iTipoDestino
    objRequisicaoCompras.lFornCliDestino = tRequisicaoCompras.lFornCliDestino
    objRequisicaoCompras.iFilialDestino = tRequisicaoCompras.iFilialDestino
    objRequisicaoCompras.lObservacao = tRequisicaoCompras.lObservacao
    objRequisicaoCompras.iTipoTributacao = tRequisicaoCompras.iTipoTributacao
    objRequisicaoCompras.sOPCodigo = tRequisicaoCompras.sOPCodigo
    objRequisicaoCompras.lNumIntDocItemOP = tRequisicaoCompras.lNumIntDocItemOP
    
    objRequisicaoCompras.sUsuReg = tRequisicaoCompras.sUsuReg
    objRequisicaoCompras.sUsuRegAprov = tRequisicaoCompras.sUsuRegAprov
    objRequisicaoCompras.sUsuRegBaixa = tRequisicaoCompras.sUsuRegBaixa
    objRequisicaoCompras.sUsuRegEnvio = tRequisicaoCompras.sUsuRegEnvio
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    RequisicaoComprasTodas_Le = SUCESSO

    Exit Function

Erro_RequisicaoComprasTodas_Le:

    RequisicaoComprasTodas_Le = gErr

    Select Case gErr

        Case 66693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66694, 66695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66696 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147699)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemRCItemPC_Le(ByVal lNumItemRC As Long, ByVal colItemPedido As Collection) As Long
'Lê todos os itens de Pedido de compras (baixados ou não) relacionados ao item de Requisição de compras passado

Dim lErro As Long
Dim lComando As Long
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim tItemPedCompraInfo As typeItemPedidoCompraInfo

On Error GoTo Erro_ItemRCItemPC_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66702
        
    tItemPedCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    
    'Pesquisa os Itens de Pedidos de Compras relacionados a Nota Fiscal a partir do NumIntDoc de NFiscal
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ItensPedCompraTodos.NumIntDoc, PedidoCompraTodos.Codigo, ItemRCItemPC.Quantidade, ItensPedCompraTodos.UM FROM PedidoCompraTodos, ItensPedCompraTodos, ItemRCItemPC WHERE PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND ItensPedCompraTodos.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ?", _
        tItemPedCompraInfo.lNumIntDoc, tItemPedCompraInfo.lPedCompra, tItemPedCompraInfo.dQuantidade, tItemPedCompraInfo.sUM, lNumItemRC)
    If lErro <> AD_SQL_SUCESSO Then gError 66703

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66704

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66706
        
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemPedCompraInfo = New ClassItemPedCompraInfo

        objItemPedCompraInfo.lPedCompra = tItemPedCompraInfo.lPedCompra
        objItemPedCompraInfo.sUM = tItemPedCompraInfo.sUM
        objItemPedCompraInfo.dQuantReceber = tItemPedCompraInfo.dQuantidade
        objItemPedCompraInfo.lNumIntDoc = tItemPedCompraInfo.lNumIntDoc
        
        'Adiciona na coleção
        colItemPedido.Add objItemPedCompraInfo

        'Busca no BD o próximo item de Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66705

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemRCItemPC_Le = SUCESSO

    Exit Function

Erro_ItemRCItemPC_Le:

    ItemRCItemPC_Le = gErr

    Select Case gErr

        Case 66702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66703, 66704, 66705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 66706 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147700)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItemReqCompra_Busca_PedidoCompra(iTela As Integer, ByVal objPedidoCompras As ClassPedidoCompras, ByVal objItemReqCompra As ClassItemReqCompras) As Long
'Busca no BD relacionamento entre as tabelas ItensPedCompra, ItemRCItemPC e ItemReqCompras a partir do objItemReqCompras passado
'retorna a tela de Pedido de Compras que será chamada : Pedidos Associados, Enviado, novo, Baixado

Dim iIndice As Integer
Dim alComando(0 To 5) As Long
Dim lErro As Long
Dim lPedCompra As Long
Dim iTotRel As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemReqCompra_Busca_PedidoCompra

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 54336
    Next

    iTotRel = 0

    'Pesquisa no BD relação entre o Item da linha selecionada e um Ítem Pedido de Compras
    lErro = Comando_Executar(alComando(0), "SELECT ItensPedCompraN.PedCompra FROM ItensPedCompraN, ItemRCItemPC WHERE ItemRCItemPC.ItemPC = ItensPedCompraN.NumIntDoc AND ItemRCItemPC.ItemRC = ?", _
        lPedCompra, objItemReqCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 54389

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54382

    objPedidoCompras.lNumIntDoc = lPedCompra

    'Enquanto encontrar relacinamento
    Do While lErro <> AD_SQL_SEM_DADOS
        'Incrementa o total de relacionamentos
        iTotRel = iTotRel + 1

        'Busca o próximo relacionamento
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54362
    Loop

    'Se o total de relacionamentos foi igual a 1
    If iTotRel = 1 Then

        'Pesquisa no BD se esse Pedido de Compra é não Baixado
        lErro = Comando_Executar(alComando(1), "SELECT ItensPedCompra.NumIntDoc FROM ItemRCItemPC, ItensPedCompra, PedidoCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ?", lNumIntDoc, objItemReqCompra.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 54361

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54360

        'Se encontrar
        If lErro = AD_SQL_SUCESSO Then
            'Seta iTela como Pedido Compras Enviado
            iTela = 0
        'Se não encontrar
        Else

            'Pesquisa no BD se esse Pedido de Compra é Baixado
            lErro = Comando_Executar(alComando(2), "SELECT PedidoCompraBaixado.NumIntDoc FROM PedidoCompraBaixado, ItensPedCompraBaixados, ItemRCItemPC WHERE PedidoCompraBaixado.NumIntDoc = ItensPedCompraBaixados.PedCompra AND ItensPedCompraBaixados.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ?", lNumIntDoc, objItemReqCompra.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then Error 54408

            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54359

            'Se encontrou
            If lErro = AD_SQL_SUCESSO Then
                'Seta iTela como Pedido Compra Baixado
                iTela = 1
            'Senão é um novo Pedido
            Else
                'Seta iTela como Pedido Compra Novo
                iTela = 2
            End If

        End If

    'Se o total de relacionamentos foi igual a 0 --> Erro
    ElseIf iTotRel = 0 Then Error 54403

    'Se o total de relacionamentos foi maior que 1
    Else
        'Seta iTela como Pedidos Associados
        iTela = 3
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

    ItemReqCompra_Busca_PedidoCompra = SUCESSO

Erro_ItemReqCompra_Busca_PedidoCompra:

    ItemReqCompra_Busca_PedidoCompra = Err
    
    Select Case Err

        Case 54336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54359, 54408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", Err)

        Case 54360, 54361, 54362, 54382, 54389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", Err)
        
        Case 54403 'Nenhum Pedido Vinculado ao ItemRC

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147701)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RequisicaoCompras_Le(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoCompras a Requisição passada por objRequisiçãoCompras

Dim lErro As Long
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long

On Error GoTo Erro_RequisicaoCompras_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54394

    With tRequisicaoCompras

        .sCcl = String(STRING_CCL, 0)
        .sDigitador = String(STRING_USUARIO, 0)
        .sOPCodigo = String(STRING_OPCODIGO, 0)
        .sUsuReg = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
        
        'Lê os dados da Requisição da tabela RequisiçãoCompras
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, NumIntDocItemOP, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompra WHERE NumIntDoc = ?", _
            .lNumIntDoc, .iFilialEmpresa, .lCodigo, .dtData, .dtDataEnvio, .dtDataLimite, .lUrgente, .lRequisitante, .sDigitador, .sCcl, .iFilialCompra, .iTipoDestino, .lFornCliDestino, .iFilialDestino, .lObservacao, .iTipoTributacao, .sOPCodigo, .lNumIntDocItemOP, .sUsuReg, .sUsuRegAprov, .sUsuRegBaixa, .sUsuRegEnvio, objRequisicaoCompras.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 54395

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54396

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 54397

        'Preenche o objRequisiçãoCompras
        objRequisicaoCompras.lNumIntDoc = .lNumIntDoc
        objRequisicaoCompras.iFilialEmpresa = .iFilialEmpresa
        objRequisicaoCompras.lCodigo = .lCodigo
        objRequisicaoCompras.dtData = .dtData
        objRequisicaoCompras.dtDataEnvio = .dtDataEnvio
        objRequisicaoCompras.dtDataLimite = .dtDataLimite
        objRequisicaoCompras.lUrgente = .lUrgente
        objRequisicaoCompras.lRequisitante = .lRequisitante
''''        objRequisicaoCompras.sDigitador = .sDigitador
        objRequisicaoCompras.sCcl = .sCcl
        objRequisicaoCompras.iFilialCompra = .iFilialCompra
        objRequisicaoCompras.iTipoDestino = .iTipoDestino
        objRequisicaoCompras.lFornCliDestino = .lFornCliDestino
        objRequisicaoCompras.iFilialDestino = .iFilialDestino
        objRequisicaoCompras.lObservacao = .lObservacao
        objRequisicaoCompras.iTipoTributacao = .iTipoTributacao
        objRequisicaoCompras.sOPCodigo = .sOPCodigo
        objRequisicaoCompras.lNumIntDocItemOP = .lNumIntDocItemOP
        
        objRequisicaoCompras.sUsuReg = .sUsuReg
        objRequisicaoCompras.sUsuRegAprov = .sUsuRegAprov
        objRequisicaoCompras.sUsuRegBaixa = .sUsuRegBaixa
        objRequisicaoCompras.sUsuRegEnvio = .sUsuRegEnvio
        
    End With

    Call Comando_Fechar(lComando)

    RequisicaoCompras_Le = SUCESSO

    Exit Function

Erro_RequisicaoCompras_Le:

    RequisicaoCompras_Le = Err

    Select Case Err

        Case 54394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54395, 54396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54397 'Tratado na rotina de chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147702)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensReqCompras_Le_ReqCompras(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras a partir do código passado em objRequisiçãoCompras

Dim lErro As Long
Dim tItemReqCompras As typeItemReqCompra
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long

On Error GoTo Erro_ItensReqCompras_Le_ReqCompras

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError (65600)

    With tItemReqCompras

        .sDescProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabil = String(STRING_BUFFER_MAX_TEXTO, 0)

        'Lê o ítem da Requisição na tabela ItensReqCompras
        lErro = Comando_Executar(lComando, "SELECT ItensReqCompra.NumIntDoc, ReqCompra, Produto, DescProduto, ItensReqCompra.Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, ItensReqCompra.Ccl, Almoxarifado, ContaContabil, ItensReqCompra.TipoTributacao, ItensReqCompra.Observacao, Fornecedor, ItensReqCompra.Filial, Exclusivo FROM ItensReqCompra, RequisicaoCompra WHERE ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.Codigo = ? AND RequisicaoCompra.FilialEmpresa = ?  ORDER BY ItensReqCompra.NumIntDoc", _
            .lNumIntDoc, .lReqCompra, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iTipoTributacao, .lObservacao, .lFornecedor, .iFilial, .iExclusivo, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 65601

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65602

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 65603

        'Equanto encontrar
        Do While lErro = AD_SQL_SUCESSO

            Set objItemReqCompras = New ClassItemReqCompras

            'Preenche o objItemReqCompras
            objItemReqCompras.lNumIntDoc = .lNumIntDoc
            objItemReqCompras.lReqCompra = .lReqCompra
            objItemReqCompras.sProduto = .sProduto
            objItemReqCompras.sDescProduto = .sDescProduto
            objItemReqCompras.iStatus = .iStatus
            objItemReqCompras.dQuantidade = .dQuantidade
            objItemReqCompras.dQuantPedida = .dQuantPedida
            objItemReqCompras.dQuantRecebida = .dQuantRecebida
            objItemReqCompras.dQuantCancelada = .dQuantCancelada
            objItemReqCompras.sUM = .sUM
            objItemReqCompras.sCcl = .sCcl
            objItemReqCompras.iAlmoxarifado = .iAlmoxarifado
            objItemReqCompras.sContaContabil = .sContaContabil
            objItemReqCompras.iTipoTributacao = .iTipoTributacao
            objItemReqCompras.lObservacao = .lObservacao
            objItemReqCompras.lFornecedor = .lFornecedor
            objItemReqCompras.iFilial = .iFilial
            objItemReqCompras.iExclusivo = .iExclusivo
            objItemReqCompras.iTipoTributacao = .iTipoTributacao
            
            'Adiciona o objItemReqCompras na coleção colItens do objRequisiçãoCompras
            objRequisicaoCompras.colItens.Add objItemReqCompras

            'Busca no BD o próximo Ítem de Requisição de compras
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65604

        Loop

    End With

    Call Comando_Fechar(lComando)

    ItensReqCompras_Le_ReqCompras = SUCESSO

    Exit Function

Erro_ItensReqCompras_Le_ReqCompras:

    ItensReqCompras_Le_ReqCompras = gErr

    Select Case gErr

        Case 65603

        Case 65600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65601, 65602, 65604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147703)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensReqModelo_Le(ByVal objReqModelo As ClassRequisicaoModelo) As Long
'Lê os Itens de uma Requisição Modelo a partir do Código dessa Requisição

Dim lErro As Long
Dim tItemReqModelo As typeItemReqModelo
Dim objItemReqModelo As ClassItemReqModelo
Dim lComando As Long
Dim lObservacao As Long

On Error GoTo Erro_ItensReqModelo_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61530

    With tItemReqModelo

        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)

        'Tenta ler os itens da Requisição Modelo a partir do código da Requisição
        lErro = Comando_Executar(lComando, "SELECT ItensReqModelo.NumIntDoc, Produto, DescProduto, Quantidade, UM, ItensReqModelo.Ccl, Almoxarifado, ContaContabil, ItensReqModelo.Observacao, ItensReqModelo.Fornecedor, ItensReqModelo.Filial, Exclusivo, ItensReqModelo.TipoTributacao FROM ItensReqModelo, RequisicaoModelo WHERE ItensReqModelo.ReqModelo = RequisicaoModelo.NumIntDoc AND RequisicaoModelo.Codigo = ? AND RequisicaoModelo.FilialEmpresa = ? ORDER BY ItensReqModelo.NumIntDoc", _
        .lNumIntDoc, .sProduto, .sDescProduto, .dQuantidade, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, lObservacao, .lFornecedor, .iFilial, .iExclusivo, .iTipoTributacao, objReqModelo.lCodigo, objReqModelo.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 61531

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61532

        'Enquanto encontrar itens
        Do While lErro <> AD_SQL_SEM_DADOS

            Set objItemReqModelo = New ClassItemReqModelo

            objItemReqModelo.dQuantidade = .dQuantidade
            objItemReqModelo.iAlmoxarifado = .iAlmoxarifado
            objItemReqModelo.iExclusivo = .iExclusivo
            objItemReqModelo.iFilial = .iFilial
            objItemReqModelo.lFornecedor = .lFornecedor
            objItemReqModelo.lNumIntDoc = .lNumIntDoc
            objItemReqModelo.sCcl = .sCcl
            objItemReqModelo.sContaContabil = .sContaContabil
            objItemReqModelo.sDescProduto = .sDescProduto
            objItemReqModelo.lObservacao = lObservacao
            objItemReqModelo.sProduto = .sProduto
            objItemReqModelo.sUM = .sUM
            objItemReqModelo.iTipoTributacao = .iTipoTributacao
            
            'Adiciona Itens na coleção
            objReqModelo.colItensReqModelo.Add objItemReqModelo

            'Busca próximo item
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61533

        Loop

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    ItensReqModelo_Le = SUCESSO

    Exit Function

Erro_ItensReqModelo_Le:

    ItensReqModelo_Le = Err

    Select Case Err

        Case 61530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61531, 61532, 61533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQMODELO", Err, objReqModelo.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147704)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicaoModelo_Le(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Lê dados da RequisiçãoModelo a partir do Número Interno passado em objRequisiçãoModelo

Dim lErro As Long
Dim tRequisicaoModelo As typeRequisicaoModelo
Dim lComando As Long

On Error GoTo Erro_RequisicaoModelo_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61575

    With tRequisicaoModelo

        .sCcl = String(STRING_CCL, 0)
        .sDescricao = String(STRING_DESCRICAO_REQMODELO, 0)
    
        'Lê os dados da Requisição Modelo a partir do NumIntDoc passado em objRequisicaoModelo
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Descricao, Urgente, Requisitante, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao,TipoTributacao FROM RequisicaoModelo WHERE NumIntDoc = ?", _
            .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sDescricao, .iUrgente, .lRequisitante, .sCcl, .iFilialCompra, .iTipoDestido, .lFornCliDestino, .iFilialDestino, .lObservacao, .iTipoTributacao, objRequisicaoModelo.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 61576

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61577

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 61578

        'Preenche o objRequisiçãoModelo
        objRequisicaoModelo.lNumIntDoc = .lNumIntDoc
        objRequisicaoModelo.iFilialEmpresa = .iFilialEmpresa
        objRequisicaoModelo.lCodigo = .lCodigo
        objRequisicaoModelo.sDescricao = .sDescricao
        objRequisicaoModelo.iUrgente = .iUrgente
        objRequisicaoModelo.lRequisitante = .lRequisitante
        objRequisicaoModelo.sCcl = .sCcl
        objRequisicaoModelo.iFilialCompra = .iFilialCompra
        objRequisicaoModelo.iTipoDestido = .iTipoDestido
        objRequisicaoModelo.lFornCliDestino = .lFornCliDestino
        objRequisicaoModelo.iFilialDestino = .iFilialDestino
        objRequisicaoModelo.lObservacao = .lObservacao
        objRequisicaoModelo.iTipoTributacao = .iTipoTributacao

    End With

    Call Comando_Fechar(lComando)

    RequisicaoModelo_Le = SUCESSO

    Exit Function

Erro_RequisicaoModelo_Le:

    RequisicaoModelo_Le = Err

    Select Case Err

        Case 61575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61576, 61577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO2", Err, objRequisicaoModelo.lNumIntDoc)

        Case 61578 'Requisição Modelo não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147705)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicaoModelo_Le_Codigo(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Lê em RequisiçãoModelo a partir do Código passado em objRequisiçãoModelo

Dim lErro As Long
Dim tRequisicaoModelo As typeRequisicaoModelo
Dim lComando As Long

On Error GoTo Erro_RequisicaoModelo_Le_codigo

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61575

    With tRequisicaoModelo

        .sCcl = String(STRING_CCL, 0)
        .sDescricao = String(STRING_DESCRICAO_REQMODELO, 0)

        'Lê os dados da Requisição Modelo a partir do Código passado em objRequisicaoModelo
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Descricao, Urgente, Requisitante, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao FROM RequisicaoModelo WHERE Codigo = ?", _
            .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sDescricao, .iUrgente, .lRequisitante, .sCcl, .iFilialCompra, .iTipoDestido, .lFornCliDestino, .iFilialDestino, .lObservacao, .iTipoTributacao, objRequisicaoModelo.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 61505

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61506

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 61508

        'Preenche o objRequisiçãoModelo
        objRequisicaoModelo.lNumIntDoc = .lNumIntDoc
        objRequisicaoModelo.iFilialEmpresa = .iFilialEmpresa
        objRequisicaoModelo.lCodigo = .lCodigo
        objRequisicaoModelo.sDescricao = .sDescricao
        objRequisicaoModelo.iUrgente = .iUrgente
        objRequisicaoModelo.lRequisitante = .lRequisitante
        objRequisicaoModelo.sCcl = .sCcl
        objRequisicaoModelo.iFilialCompra = .iFilialCompra
        objRequisicaoModelo.iTipoDestido = .iTipoDestido
        objRequisicaoModelo.lFornCliDestino = .lFornCliDestino
        objRequisicaoModelo.iFilialDestino = .iFilialDestino
        objRequisicaoModelo.lObservacao = .lObservacao
        objRequisicaoModelo.iTipoTributacao = .iTipoTributacao


    End With

    Call Comando_Fechar(lComando)

    RequisicaoModelo_Le_Codigo = SUCESSO

    Exit Function

Erro_RequisicaoModelo_Le_codigo:

    RequisicaoModelo_Le_Codigo = Err

    Select Case Err

        Case 61575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61505, 61506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)

        Case 61508 'Tratado na rotina de chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147706)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Requisitante_Valida_Exclusao(ByVal objRequisitante As ClassRequisitante, alComando() As Long) As Long
'Verifica se o Requisitante a ser excluído está vinculado a alguma outra tabela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Requisitante_Valida_Exclusao

    'verifica se Requisitante esta relacionado com alguma Requisicao de Compra
    lErro = Comando_Executar(alComando(1), "SELECT Requisitante FROM RequisicaoCompra WHERE Requisitante =?", lCodigo, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49079

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49098
    
    'O requisitante está relacionado com Requisicao de Compra
    If lErro = AD_SQL_SUCESSO Then Error 49099
    
    'verifica se Requisitante esta relacionado com alguma Requisicao de Compra Baixada
    lErro = Comando_Executar(alComando(1), "SELECT Requisitante FROM RequisicaoCompraBaixada WHERE Requisitante =?", lCodigo, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49291

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49292
    
    'O requisitante está relacionado com Requisicao de Compra Baixada
    If lErro = AD_SQL_SUCESSO Then Error 49293
    
    'verifica se Requisitante esta relacionado com alguma Requisicao Modelo
    lErro = Comando_Executar(alComando(1), "SELECT Requisitante FROM RequisicaoModelo WHERE Requisitante =?", lCodigo, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49294

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49295
    
    'O requisitante está relacionado com Requisicao Modelo
    If lErro = AD_SQL_SUCESSO Then Error 49296

    Requisitante_Valida_Exclusao = SUCESSO
    
    Exit Function
    
Erro_Requisitante_Valida_Exclusao:

    Requisitante_Valida_Exclusao = Err
    
    Select Case Err
    
        Case 49079, 49098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA1", Err, objRequisitante.lCodigo)
            
        Case 49099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_VINCULADO_REQUISICAOCOMPRA", Err, objRequisitante.lCodigo)
        
        Case 49291, 49292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRABAIXADA1", Err, objRequisitante.lCodigo)
        
        Case 49293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_VINCULADO_REQCOMPRABAIXADA", Err, objRequisitante.lCodigo)
            
        Case 49294, 49295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO1", Err, objRequisitante.lCodigo)
        
        Case 49296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_VINCULADO_REQCOMPRAMODELO", Err, objRequisitante.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147707)

    End Select
    
    Exit Function

End Function

Function TipoDeBloqueioPC_Le(ByVal objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Le todos os dados da tabela TipoDeBloqueioPC, correspondente a chave passada por parametro

Dim lComando As Long
Dim lErro As Long
Dim tTipoBloqueioPC As typeTipoBloqueioPC

On Error GoTo Erro_TipoDeBloqueioPC_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49140

    tTipoBloqueioPC.sDescricao = String(STRING_TIPODEBLOQUEIOPC_DESCRICAO, 0)
    tTipoBloqueioPC.sNomeReduzido = String(STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, 0)

    'Pega os dados da tabela TiposDeBloqueioPC
    lErro = Comando_Executar(lComando, "SELECT Descricao,NomeReduzido,Codigo FROM TiposDeBloqueioPC WHERE Codigo = ? ", tTipoBloqueioPC.sDescricao, tTipoBloqueioPC.sNomeReduzido, tTipoBloqueioPC.iCodigo, objTipoDeBloqueioPC.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49141

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19142

    'Tipo de Bloqueio PC nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 49143

    'carrega os dados em objTipoDeBloqueioPC
    objTipoDeBloqueioPC.sDescricao = tTipoBloqueioPC.sDescricao
    objTipoDeBloqueioPC.sNomeReduzido = tTipoBloqueioPC.sNomeReduzido
    objTipoDeBloqueioPC.iCodigo = tTipoBloqueioPC.iCodigo

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    TipoDeBloqueioPC_Le = SUCESSO

    Exit Function

Erro_TipoDeBloqueioPC_Le:

    TipoDeBloqueioPC_Le = Err

    Select Case Err

        Case 49140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49141, 49142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49143

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147708)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function BaixaReqCompras_Obter_Residuais(ByVal objBaixaReqCompras As ClassBaixaReqCompras) As Long
'Verifica quais Requisicoes de Compras não são residuais e retira- as de objBaixaReqCompras.colReqComprasInfo
' Para trazer as residuais:
' 1) olhar primeiro em produto se o resíduo está preenchido (<> -1)
' 2) Se em produto não estiver preenchido olhar se está em TipoDeproduto
' 3) Se não estiver preenchido em TP Produto olhar em Compras config
'* O Valor zero é um valor válido para cálculo de resíduo, logo,
'só qdo o resíduo estiver com -1 é que significa que ele não foi preenchido.

Dim lErro As Long
Dim objReqComprasInfo As New ClassReqComprasInfo
Dim iIndice As Integer
Dim objItemReqCompras As New ClassItemReqCompras
Dim objRequisicaoCompras As New ClassReqComprasInfo
Dim iResidual As Integer
Dim objProduto As New ClassProduto
Dim objRequisicaoCompras1 As New ClassRequisicaoCompras
Dim iIndice2 As Integer
Dim dResiduo As Double
Dim dMinPercRec  As Double
Dim iNumItens As Integer

On Error GoTo Erro_BaixaReqCompras_Obter_Residuais

    'Percorre os elementos da colecao ReqComprasInfo
    For iIndice = objBaixaReqCompras.colReqComprasInfo.Count To 1 Step -1
    
        Set objRequisicaoCompras = objBaixaReqCompras.colReqComprasInfo(iIndice)
        
        objRequisicaoCompras1.lNumIntDoc = objRequisicaoCompras.lNumIntReq
        
        Set objRequisicaoCompras1.colItens = New Collection
        
        'Lê os itens da Requisicao de Compras
        lErro = CF("ItensReqCompras_Le", objRequisicaoCompras1)
        If lErro <> SUCESSO And lErro <> 54377 Then Error 63427
        
        'Se não encontrou ==> erro
        If lErro = 54377 Then Error 63428
        
        iNumItens = objRequisicaoCompras1.colItens.Count
        
        'Percorre os elementos de colItens, que contem os itens de uma Requisicao de Compra
        For iIndice2 = objRequisicaoCompras1.colItens.Count To 1 Step -1
            
            Set objItemReqCompras = objRequisicaoCompras1.colItens(iIndice2)
            
            objProduto.sCodigo = objItemReqCompras.sProduto
            
            'Lê o valor do residuo para o produto passado como parâmetro
            lErro = CF("Produto_Le_Residual", objProduto, dResiduo)
            If lErro <> SUCESSO Then Error 63425
            
            If objItemReqCompras.dQuantidade <> 0 Then dMinPercRec = (objItemReqCompras.dQuantRecebida + objItemReqCompras.dQuantCancelada) / objItemReqCompras.dQuantidade
                
            'Verifica se é um Item Residual
            If dResiduo + dMinPercRec >= 1 Then
                iResidual = iResidual + 1
            Else
                Exit For
            End If

        Next
        
        'Se a requisição não é residual
        If iResidual <> iNumItens Then

            'remove a requisição de Compras da colecao
            objBaixaReqCompras.colReqComprasInfo.Remove (iIndice)

        End If

        iResidual = 0
        
    Next
    
    BaixaReqCompras_Obter_Residuais = SUCESSO
    
    Exit Function
    
Erro_BaixaReqCompras_Obter_Residuais:

    BaixaReqCompras_Obter_Residuais = Err
    
    Select Case Err
        
        Case 63425, 63427
            'Erros tratados nas rotinas chamadas
                
        Case 63428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO", Err, objItemReqCompras.lNumIntDoc, objRequisicaoCompras1.lCodigo)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147709)

    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Le_CotacaoItemConcorrencia(ByVal objCotacaoItemConc As ClassCotacaoItemConc, ByVal objConcorrencia As ClassConcorrencia) As Long
''Lê a Concorrencia a partir do NumIntDoc de CotacaoItemConcorrencia

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoItemConc As typeCotacaoItemConcorrencia
Dim lNumIntDoc As Long
Dim lConcorrencia As Long
Dim tConcorrencia As typeConcorrencia
Dim lItemConcorrencia As Long

On Error GoTo Erro_ItensPedCompra_Le_CotacaoItemConcorrencia

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 74854

    'Busca em CotacaoItemConcorrenciaBaixado, a partir do NumIntDoc informado
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,ItemConcorrencia FROM CotacaoItemConcorrenciaTodos WHERE NumIntDoc=?", tCotacaoItemConc.lNumIntDoc, lItemConcorrencia, objCotacaoItemConc.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 74855
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74856
    
    'Busca o item de concorrencia baixado a partir do NumIntDoc informado
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Concorrencia FROM ItensConcorrenciaBaixados WHERE NumIntDoc=?", lNumIntDoc, lConcorrencia, lItemConcorrencia)
    If lErro <> AD_SQL_SUCESSO Then gError 74857
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74858
    
    'Busca no Bd a ConcorrenciaBaixada cujo NumIntDoc foi informado
    objConcorrencia.lNumIntDoc = lConcorrencia
    lErro = ConcorrenciaBaixada_Le1(objConcorrencia)
    If lErro <> SUCESSO And lErro <> 81624 Then gError 74859
    If lErro <> SUCESSO Then gError 86154
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPedCompra_Le_CotacaoItemConcorrencia = SUCESSO
    
    Exit Function
    
Erro_ItensPedCompra_Le_CotacaoItemConcorrencia:

    ItensPedCompra_Le_CotacaoItemConcorrencia = gErr
    
    Select Case gErr
    
        Case 74854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 74855, 74856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIABAIXADA", gErr)
            
        Case 74857, 74858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCONCORRENCIABAIXADAS", gErr)
            
        Case 74859
        
        Case 86154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147710)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RequisicaoComprasTodas_Le_Numero(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoComprasTodas a Requisição(baixada ou não) passada por objRequisiçãoCompras com Código e FlialEmpresa

Dim lErro As Long
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long

On Error GoTo Erro_RequisicaoComprasTodas_Le_Numero

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76059

    With tRequisicaoCompras

        .sCcl = String(STRING_CCL, 0)
        .sOPCodigo = String(STRING_OPCODIGO, 0)
        .sDigitador = String(STRING_USUARIO_CODIGO, 0)
        .sUsuReg = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
        .sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
        
        'Lê os dados da Requisição da tabela RequisiçãoCompras
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataBaixa, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, NumIntDocItemOP, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompraTodas WHERE Codigo = ? AND FilialEmpresa = ?", _
            .lNumIntDoc, .iFilialEmpresa, .lCodigo, .dtData, .dtDataEnvio, .dtDataBaixa, .dtDataLimite, .lUrgente, .lRequisitante, .sDigitador, .sCcl, .iFilialCompra, .iTipoDestino, .lFornCliDestino, .iFilialDestino, .lObservacao, .iTipoTributacao, .sOPCodigo, .lNumIntDocItemOP, .sUsuReg, .sUsuRegAprov, .sUsuRegBaixa, .sUsuRegEnvio, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 76060

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76061

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 76062

        'Preenche o objRequisiçãoCompras
        objRequisicaoCompras.lNumIntDoc = .lNumIntDoc
        objRequisicaoCompras.iFilialEmpresa = .iFilialEmpresa
        objRequisicaoCompras.lCodigo = .lCodigo
        objRequisicaoCompras.dtData = .dtData
        objRequisicaoCompras.dtDataEnvio = .dtDataEnvio
        objRequisicaoCompras.dtDataBaixa = .dtDataBaixa
        objRequisicaoCompras.dtDataLimite = .dtDataLimite
        objRequisicaoCompras.lUrgente = .lUrgente
        objRequisicaoCompras.lRequisitante = .lRequisitante
''''        objRequisicaoCompras.sDigitador = .sDigitador
        objRequisicaoCompras.sCcl = .sCcl
        objRequisicaoCompras.iFilialCompra = .iFilialCompra
        objRequisicaoCompras.iTipoDestino = .iTipoDestino
        objRequisicaoCompras.lFornCliDestino = .lFornCliDestino
        objRequisicaoCompras.iFilialDestino = .iFilialDestino
        objRequisicaoCompras.lObservacao = .lObservacao
        objRequisicaoCompras.sOPCodigo = .sOPCodigo
        objRequisicaoCompras.lNumIntDocItemOP = .lNumIntDocItemOP
        
        objRequisicaoCompras.sUsuReg = .sUsuReg
        objRequisicaoCompras.sUsuRegAprov = .sUsuRegAprov
        objRequisicaoCompras.sUsuRegBaixa = .sUsuRegBaixa
        objRequisicaoCompras.sUsuRegEnvio = .sUsuRegEnvio
        
    End With

    Call Comando_Fechar(lComando)

    RequisicaoComprasTodas_Le_Numero = SUCESSO

    Exit Function

Erro_RequisicaoComprasTodas_Le_Numero:

    RequisicaoComprasTodas_Le_Numero = gErr

    Select Case gErr

        Case 76059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 76060, 76061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case 76062 'Tratado na rotina de chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147711)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TP_Comprador_Le(ByVal objCompradorMaskEdBox As Object, ByVal objComprador As ClassComprador, Optional ByVal iCria As Integer = 1) As Long
'Lê o Comprador com Código ou NomeRed em objCompradorMaskEdBox.Text
'Devolve em objComprador. Coloca código-NomeReduzido no .Text

Dim sComprador As String
Dim iCodigo As Long
Dim Comprador As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuario As New ClassUsuario

On Error GoTo TP_Comprador_Le

    Set Comprador = objCompradorMaskEdBox
    sComprador = Trim(Comprador.Text)

    'Tenta extrair código de sComprador
    iCodigo = Codigo_Extrai(sComprador)

    'Se é do tipo código
    If iCodigo > 0 Then

        objComprador.iCodigo = iCodigo
        objComprador.iFilialEmpresa = giFilialEmpresa

        'Lê o comprador
        lErro = CF("Comprador_Le", objComprador)
        If lErro <> SUCESSO And lErro <> 50064 Then gError 68618
        If lErro <> SUCESSO Then gError 68619

        objUsuario.sCodUsuario = objComprador.sCodUsuario
        
        'Lê o usuário
        lErro = CF("Usuario_Le", objUsuario)
        If lErro <> SUCESSO And lErro <> 36347 Then gError 72550
        If lErro <> SUCESSO Then gError 72551
        
        Comprador.Text = CStr(objComprador.iCodigo) & SEPARADOR & objUsuario.sNomeReduzido

    Else  'Se é do tipo String

         objUsuario.sNomeReduzido = sComprador

         lErro = CF("Usuario_Le_NomeRed", objUsuario)
         If lErro <> SUCESSO And lErro <> 57269 Then gError 68620
         If lErro = 57269 Then gError 68621

        objComprador.sCodUsuario = objUsuario.sCodUsuario
        
        lErro = CF("Comprador_Le_Usuario", objComprador)
        If lErro <> SUCESSO And lErro <> 50059 Then gError 68622
        If lErro <> SUCESSO Then gError 68623

        Comprador.Text = CStr(objComprador.iCodigo) & SEPARADOR & sComprador


    End If

    TP_Comprador_Le = SUCESSO

    Exit Function

TP_Comprador_Le:

    TP_Comprador_Le = gErr

    Select Case gErr

        Case 68618, 68620, 68622, 72550
        
        Case 68619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO", gErr, objComprador.iCodigo)
            
        Case 68621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)

        Case 68623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR2", gErr, objUsuario.sCodUsuario)
            
        Case 72551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147712)

    End Select
    
    Exit Function

End Function

'Janaina - 19/07 - Estava em PedidoCotacaoOcx
'???? Transferir para ClassCOMSelect
Function ItemPedCotacao_Le_RequisicaoCompra(ByVal objItemPedCotacao As ClassItemPedCotacao, ByVal colReqCompras As Collection, ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPedCompra As New ClassItemPedCompra
Dim objReqCompras As New ClassRequisicaoCompras
Dim objItemRC As New ClassItemReqCompras
Dim objItemPC As New ClassItemPedCompra
Dim dQuantidadeCotar As Double
Dim lNumIntDocCotProduto As Long
Dim lCotacaoProduto As Long
Dim lItemReqCompra As Long
Dim colPedCompraExclu As New Collection
Dim colPedCompraGeral As New Collection
Dim tItemRC As typeItemReqCompra
Dim tReqCompras As typeRequisicaoCompras
Dim lCotacaoProdutoItemPC As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemPedCotacao_Le_RequisicaoCompra

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76151

    tItemRC.sProduto = String(STRING_PRODUTO, 0)
    tItemRC.sUM = String(STRING_UM_SIGLA, 0)
    
    'Se o produto tem fornecedor filial amarrados
    If objItemPedCotacao.iExclusivo = MARCADO Then
        'Busca o Item de Pedido de Cotacao no BD
        lErro = Comando_Executar(lComando, "SELECT ItemPedCotacao.NumIntDoc FROM ItemPedCotacao,PedidoCotacao,CotacaoProduto WHERE PedidoCotacao.NumIntDoc=ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND CotacaoProduto.Fornecedor > 0 AND CotacaoProduto.Filial > 0 AND CotacaoProduto.Produto=? AND PedidoCotacao.Codigo=? AND PedidoCotacao.FilialEmpresa=?", lNumIntDoc, objItemPedCotacao.sProduto, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    Else
        'Busca o Item de Pedido de Cotacao no BD
        lErro = Comando_Executar(lComando, "SELECT ItemPedCotacao.NumIntDoc FROM ItemPedCotacao,PedidoCotacao,CotacaoProduto WHERE PedidoCotacao.NumIntDoc=ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND CotacaoProduto.Fornecedor = 0 AND CotacaoProduto.Filial = 0 AND CotacaoProduto.Produto=? AND PedidoCotacao.Codigo=? AND PedidoCotacao.FilialEmpresa=?", lNumIntDoc, objItemPedCotacao.sProduto, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 76152
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76153
    If lErro = AD_SQL_SEM_DADOS Then gError 76154
    
    objItemPedCotacao.lNumIntDoc = lNumIntDoc
    
    'Lê os ItensReqCompra associados ao ItemPedCotacao
    lErro = Comando_Executar(lComando, "SELECT ItensReqCompra.UM,ItensReqCompra.NumIntDoc,RequisicaoCompra.TipoDestino,RequisicaoCompra.FornCliDestino,RequisicaoCompra.FilialDestino,ItensReqCompra.Exclusivo,CotacaoProdutoItemRC.QuantidadeCotar,ItemPedCotacao.CotacaoProduto,CotacaoProduto.NumIntDoc,CotacaoProdutoItemRC.CotacaoProduto,CotacaoProdutoItemRC.ItemReqCompra,ItensReqCompra.ReqCompra,RequisicaoCompra.NumIntDoc,ItensReqCompra.Produto FROM ItemPedCotacao,CotacaoProduto,CotacaoProdutoItemRC,ItensReqCompra,RequisicaoCompra WHERE ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND CotacaoProduto.NumIntDoc=CotacaoProdutoItemRC.CotacaoProduto AND CotacaoProdutoItemRC.ItemReqCompra=ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra=RequisicaoCompra.NumIntDoc AND ItemPedCotacao.NumIntDoc=?", _
    tItemRC.sUM, tItemRC.lNumIntDoc, tReqCompras.iTipoDestino, tReqCompras.lFornCliDestino, tReqCompras.iFilialDestino, tItemRC.iExclusivo, dQuantidadeCotar, lCotacaoProdutoItemPC, lNumIntDocCotProduto, lCotacaoProduto, lItemReqCompra, tItemRC.lReqCompra, tReqCompras.lNumIntDoc, tItemRC.sProduto, objItemPedCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 76155
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76156
    If lErro = AD_SQL_SEM_DADOS Then gError 76158
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemRC = New ClassItemReqCompras
        Set objReqCompras = New ClassRequisicaoCompras
        
        'Preenche objItemRC
        objItemRC.iExclusivo = tItemRC.iExclusivo
        objItemRC.sProduto = tItemRC.sProduto
        objItemRC.dQuantNaCotacao = dQuantidadeCotar
        objItemRC.lNumIntDoc = tItemRC.lNumIntDoc
        objItemRC.sUM = tItemRC.sUM
        
        'Preenche objReqCompras
        objReqCompras.iTipoDestino = tReqCompras.iTipoDestino
        objReqCompras.lFornCliDestino = tReqCompras.lFornCliDestino
        objReqCompras.iFilialDestino = tReqCompras.iFilialDestino
        
        'Adiciona na colecao de ItensReqCompras
        objReqCompras.colItens.Add objItemRC
        
        'Adiciona na colecao de ReqCompras
        colReqCompras.Add objReqCompras
        
        'Busca a proxima ReqCompra
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76157
        
    Loop
    
    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    ItemPedCotacao_Le_RequisicaoCompra = SUCESSO
    
    Exit Function
    
Erro_ItemPedCotacao_Le_RequisicaoCompra:

    ItemPedCotacao_Le_RequisicaoCompra = gErr
    
    Select Case gErr
    
        Case 76151
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76152, 76153
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO1", gErr)
            
        Case 76154
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)
            
        Case 76155, 76156, 76157
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
            
        Case 76158
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2", gErr, tItemRC.lNumIntDoc)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147713)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    

End Function

Function Valida_Quantidade_ItemPedCotacao(ByVal objPedCotacao As ClassPedidoCotacao) As Long
'Faz a validação do ItemPedCotacao, ou seja, verifica se a quantidade
'do Item de Pedido de Cotacao está toda vinculada à Itens de Requisicao

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim tCotacaoProduto As typeCotacaoProduto
Dim lNumIntItemPedCotacao As Long
Dim dQuantCotarTotal  As Double
Dim lCotacaoProduto As Long
Dim dQuantCotar As Double
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim lNumIntDoc As Long

On Error GoTo Erro_Valida_Quantidade_ItemPedCotacao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76134

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 76147
    
    'Busca o NumIntDoc do ItemPedCotacao
    lErro = Comando_Executar(lComando2, "SELECT ItemPedCotacao.NumIntDoc FROM ItemPedCotacao,PedidoCotacao WHERE PedidoCotacao.NumIntDoc=ItemPedCotacao.PedidoCotacao AND PedidoCotacao.Codigo=? AND PedidoCotacao.FilialEmpresa=?", lNumIntDoc, objPedCotacao.lCodigo, objPedCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 76135

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76136

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 76137

    Do While lErro = AD_SQL_SUCESSO
    
        objItemPedCotacao.lNumIntDoc = lNumIntDoc
    
        'Pesquisa no BD a CotacaoProduto ligada ao ItemPedCotacao passado como parametro
        lErro = Comando_Executar(lComando, "SELECT CotacaoProduto.NumIntDoc,CotacaoProduto.Quantidade,ItemPedCotacao.CotacaoProduto FROM CotacaoProduto,ItemPedCotacao WHERE ItemPedCotacao.NumIntDoc=? AND ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc", tCotacaoProduto.lNumIntDoc, tCotacaoProduto.dQuantidade, lNumIntItemPedCotacao, objItemPedCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 76138
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76139
    
        'Se não encontrou ==> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 76140
    
        'Verifica se existe CotacaoProdutoItemRC associada ao ItemPedCotacao
        lErro = Comando_Executar(lComando, "SELECT CotacaoProduto,QuantidadeCotar FROM CotacaoProdutoItemRC WHERE CotacaoProduto=?", lCotacaoProduto, dQuantCotar, tCotacaoProduto.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 76141
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76142
    
        'Se não encontrou ==> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 76143
    
        Do While lErro = AD_SQL_SUCESSO
    
            'Soma  QuantidadeCotar
            dQuantCotarTotal = dQuantCotar + dQuantCotarTotal
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76144
        
        Loop
    
        'Verifica se a Quantidade do item do pedido de cotacao
        'é diferente da soma das quantidades cotar
        If tCotacaoProduto.dQuantidade <> dQuantCotarTotal Then gError 76145
    
        dQuantCotarTotal = 0
        dQuantCotar = 0
        
        'Busca o próximo Item de Pedido de Cotacao
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76148
        
    Loop
        
    'Fecha o comando
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)
    
    Valida_Quantidade_ItemPedCotacao = SUCESSO
    
    
    Exit Function
    
Erro_Valida_Quantidade_ItemPedCotacao:
    
    Valida_Quantidade_ItemPedCotacao = gErr
    
    Select Case gErr
    
        Case 76134, 76147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76135, 76136, 76148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO1", gErr)
            
        Case 76137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO", gErr, objPedCotacao.lCodigo)
            
        Case 76140, 76143
        
        Case 76138, 76139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)
            
        Case 76141, 76142, 76144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)
            
        Case 76145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_COTAR_MAIOR_QUANTIDADE_ITEMPEDCOTACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147714)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function




Private Sub Busca_Na_Colecao(ByVal collCodigoNome As AdmCollCodigoNome, ByVal lCodigo As Long, iPosicao As Integer)
'Busca a chave lCodigo na coleção

Dim objlCodigoNome As AdmlCodigoNome
Dim iIndice As Integer

    iPosicao = 0
    iIndice = 0
    
    'Para cada item da coleção
    For Each objlCodigoNome In collCodigoNome
        
        iIndice = iIndice + 1
        
        'Busca o item com a chave passada
        If objlCodigoNome.lCodigo = lCodigo Then
            
            iPosicao = iIndice
            Exit For
        
        End If
    
    Next
    
    Exit Sub

End Sub

Function TiposProduto_Le_Todos(ByVal colCod_DescReduzida As AdmColCodigoNome) As Long
'Le todos os Codigos e DescReduzida dos tipos de produto cadastrados

Dim lComando As Long
Dim lErro As Long
Dim tTipoDeProduto As typeTipoDeProduto

On Error GoTo Erro_TiposProduto_Le_Todos

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 75393

    tTipoDeProduto.sDescricao = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)

    'Seleciona Tipos de Produto
    lErro = Comando_Executar(lComando, "SELECT TipoDeProduto, Descricao FROM TiposDeProduto ORDER BY TipoDeProduto", tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 75394

    'Lê o primeiro Codigo e DescReduzida
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75395

    'Enquanto Encontrar o Tipo de Produto
    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o Codigo e a DescReduzida lidos na coleção
        colCod_DescReduzida.Add tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao

        'Lê o próximo Codigo e DescReduzida
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75396

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    TiposProduto_Le_Todos = SUCESSO

    Exit Function

Erro_TiposProduto_Le_Todos:

    TiposProduto_Le_Todos = gErr

    Select Case gErr

        Case 75393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO1", gErr)

        Case 75395, 75396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", gErr, tTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147715)

    End Select

    'Fecha comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemCotacao_Le(ByVal objItemCotacao As ClassItemCotacao) As Long
'Lê o item de cotação com o numintdsoc passado

Dim lErro As Long
Dim lComando As Long
Dim tItemCotacao As typeItemCotacao

On Error GoTo Erro_ItemCotacao_Le
    
    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62721
    
    'BUsca o ItemCotacao no BD
    lErro = Comando_Executar(lComando, "SELECT CondPagto,DataReferencia,PrecoUnitario,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,AliquotaIPI,AliquotaICMS,PrazoEntrega,QuantEntrega,Observacao, Moeda, Taxa FROM ItensCotacao WHERE NumIntDoc  = ? ", tItemCotacao.iCondPagto, tItemCotacao.dtDataReferencia, tItemCotacao.dPrecoUnitario, tItemCotacao.dOutrasDespesas, tItemCotacao.dValorFrete, tItemCotacao.dValorSeguro, tItemCotacao.dValorDesconto, tItemCotacao.dValorTotal, tItemCotacao.dValorIPI, tItemCotacao.dAliquotaIPI, tItemCotacao.dAliquotaICMS, tItemCotacao.iPrazoEntrega, tItemCotacao.dQuantEntrega, tItemCotacao.lObservacao, tItemCotacao.iMoeda, tItemCotacao.dTaxa, objItemCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 62722
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62723
    If lErro <> AD_SQL_SUCESSO Then Error 62724 'Não encontrou
    
    With tItemCotacao
        
        'Transfere os dados lidos para o objitemcotacao
        objItemCotacao.dAliquotaICMS = .dAliquotaICMS
        objItemCotacao.dAliquotaIPI = .dAliquotaIPI
        objItemCotacao.dOutrasDespesas = .dOutrasDespesas
        objItemCotacao.dPrecoUnitario = .dPrecoUnitario
        objItemCotacao.dQuantEntrega = .dQuantEntrega
        objItemCotacao.dtDataReferencia = .dtDataReferencia
        objItemCotacao.dValorDesconto = .dValorDesconto
        objItemCotacao.dValorFrete = .dValorFrete
        objItemCotacao.dValorIPI = .dValorIPI
        objItemCotacao.dValorSeguro = .dValorSeguro
        objItemCotacao.dValorTotal = .dValorTotal
        objItemCotacao.iCondPagto = .iCondPagto
        objItemCotacao.iPrazoEntrega = .iPrazoEntrega
        objItemCotacao.lObservacao = .lObservacao
        objItemCotacao.iMoeda = .iMoeda
        objItemCotacao.dTaxa = .dTaxa
    
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemCotacao_Le = SUCESSO
    
    Exit Function
    
Erro_ItemCotacao_Le:

    ItemCotacao_Le = Err
    
    Select Case Err
    
        Case 62721
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62722, 62723
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", Err)
        
        Case 62724
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCOTACAO_NAO_CADASTRADO", Err, objItemCotacao.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147716)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ItensReqCompra_Le_GeracaoPC(ByVal objGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Lê os itens de Requisição dos Tipos passados por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItensReqCompra As ClassItemReqCompras
Dim tItemReqCompra As typeItemReqCompra
Dim iIndice As Integer
Dim iTipo As Integer
Dim bAchou As Boolean
Dim sItemCategoria As String
Dim bItemCategoriaCerto As Boolean

On Error GoTo Erro_ItensReqCompra_Le_GeracaoPC

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62730

    'Para cada Requisicao
    For Each objReqCompra In objGeracaoPedCompraReq.colRequisicao

        'Inicializa os campos texto do BD
        tItemReqCompra.sCcl = String(STRING_CCL, 0)
        tItemReqCompra.sContaContabil = String(STRING_CONTA, 0)
        tItemReqCompra.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        tItemReqCompra.sProduto = String(STRING_PRODUTO, 0)
        tItemReqCompra.sUM = String(STRING_UM_SIGLA, 0)
        tItemReqCompra.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
        sItemCategoria = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        
        If Len(Trim(objGeracaoPedCompraReq.sCategoria)) > 0 Then
            'Busca os itens com os produtos do tipo passado
            lErro = Comando_Executar(lComando, "SELECT ProdutoCategoria.Item, ItensReqCompra.NumIntDoc, ItensReqCompra.Produto, ItensReqCompra.DescProduto, ItensReqCompra.Status, ItensReqCompra.Quantidade, ItensReqCompra.QuantPedida, ItensReqCompra.QuantRecebida, ItensReqCompra.QuantCancelada, ItensReqCompra.UM, ItensReqCompra.Ccl, ItensReqCompra.Almoxarifado, ItensReqCompra.ContaContabil, ItensReqCompra.CreditaICMS, ItensReqCompra.CreditaIPI, ItensReqCompra.Observacao, ItensReqCompra.Fornecedor, ItensReqCompra.Filial, ItensReqCompra.Exclusivo, ItensReqCompra.TipoTributacao, Produtos.Tipo, Observacao.Observacao AS Obs FROM ProdutoCategoria INNER JOIN ItensReqCompra INNER JOIN Produtos ON ItensReqCompra.Produto = Produtos.Codigo ON ProdutoCategoria.Produto = ItensReqCompra.Produto LEFT OUTER JOIN Observacao ON ItensReqCompra.Observacao = Observacao.NumInt WHERE (ItensReqCompra.ReqCompra = ?) AND ProdutoCategoria.Categoria = ? ORDER BY ItensReqCompra.NumIntDoc", _
            sItemCategoria, tItemReqCompra.lNumIntDoc, tItemReqCompra.sProduto, tItemReqCompra.sDescProduto, tItemReqCompra.iStatus, tItemReqCompra.dQuantidade, tItemReqCompra.dQuantPedida, tItemReqCompra.dQuantRecebida, tItemReqCompra.dQuantCancelada, tItemReqCompra.sUM, tItemReqCompra.sCcl, tItemReqCompra.iAlmoxarifado, tItemReqCompra.sContaContabil, tItemReqCompra.iCreditaICMS, tItemReqCompra.iCreditaIPI, tItemReqCompra.lObservacao, tItemReqCompra.lFornecedor, tItemReqCompra.iFilial, tItemReqCompra.iExclusivo, tItemReqCompra.iTipoTributacao, iTipo, tItemReqCompra.sObservacao, objReqCompra.lNumIntDoc, objGeracaoPedCompraReq.sCategoria)
        Else
            'Busca os itens com os produtos do tipo passado
            lErro = Comando_Executar(lComando, "SELECT ItensReqCompra.NumIntDoc, ItensReqCompra.Produto, ItensReqCompra.DescProduto, ItensReqCompra.Status, ItensReqCompra.Quantidade, ItensReqCompra.QuantPedida, ItensReqCompra.QuantRecebida, ItensReqCompra.QuantCancelada, ItensReqCompra.UM, ItensReqCompra.Ccl, ItensReqCompra.Almoxarifado, ItensReqCompra.ContaContabil, ItensReqCompra.CreditaICMS, ItensReqCompra.CreditaIPI, ItensReqCompra.Observacao, ItensReqCompra.Fornecedor, ItensReqCompra.Filial, ItensReqCompra.Exclusivo, ItensReqCompra.TipoTributacao, Produtos.Tipo, Observacao.Observacao AS Obs FROM ItensReqCompra INNER JOIN Produtos ON ItensReqCompra.Produto = Produtos.Codigo LEFT OUTER JOIN Observacao ON ItensReqCompra.Observacao = Observacao.NumInt WHERE (ItensReqCompra.ReqCompra = ?) ORDER BY ItensReqCompra.NumIntDoc", _
            tItemReqCompra.lNumIntDoc, tItemReqCompra.sProduto, tItemReqCompra.sDescProduto, tItemReqCompra.iStatus, tItemReqCompra.dQuantidade, tItemReqCompra.dQuantPedida, tItemReqCompra.dQuantRecebida, tItemReqCompra.dQuantCancelada, tItemReqCompra.sUM, tItemReqCompra.sCcl, tItemReqCompra.iAlmoxarifado, tItemReqCompra.sContaContabil, tItemReqCompra.iCreditaICMS, tItemReqCompra.iCreditaIPI, tItemReqCompra.lObservacao, tItemReqCompra.lFornecedor, tItemReqCompra.iFilial, tItemReqCompra.iExclusivo, tItemReqCompra.iTipoTributacao, iTipo, tItemReqCompra.sObservacao, objReqCompra.lNumIntDoc)
        End If
        
        If lErro <> AD_SQL_SUCESSO Then Error 62734

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62735
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            bItemCategoriaCerto = False
            
            'Define se será incluido na colecao
            'Caso seja da mesma categoria/item => Adiciona
            For iIndice = 1 To objGeracaoPedCompraReq.colTipoCategoria.Count
                If objGeracaoPedCompraReq.colTipoCategoria.Item(iIndice) = sItemCategoria Then
                    bItemCategoriaCerto = True
                    Exit For
                End If
            Next
            
            'Se a categoria nao foi preenchida => Adiciona quaisquer.
            If Len(Trim(objGeracaoPedCompraReq.sCategoria)) = 0 Then bItemCategoriaCerto = True
            
            If Not bItemCategoriaCerto Then
            
                    'Busca o Próximo Item
                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62737
                    
            Else
    
                bAchou = False
                For iIndice = 1 To objGeracaoPedCompraReq.colTipoProduto.Count
                    If iTipo = objGeracaoPedCompraReq.colTipoProduto(iIndice) Then
                        bAchou = True
                    End If
                Next
                
                If bAchou And (tItemReqCompra.dQuantidade - tItemReqCompra.dQuantCancelada - tItemReqCompra.dQuantPedida - tItemReqCompra.dQuantRecebida > 0) Then
                    'Preenche um objItemReq
                    Set objItensReqCompra = New ClassItemReqCompras
        
                    objItensReqCompra.dQuantCancelada = tItemReqCompra.dQuantCancelada
                    objItensReqCompra.dQuantidade = tItemReqCompra.dQuantidade
                    objItensReqCompra.dQuantNaConcorrencia = tItemReqCompra.dQuantNaConcorrencia
                    objItensReqCompra.dQuantNaCotacao = tItemReqCompra.dQuantNaCotacao
                    objItensReqCompra.dQuantNoPedido = tItemReqCompra.dQuantNoPedido
                    objItensReqCompra.dQuantNoPedidoRecebida = tItemReqCompra.dQuantNoPedidoRecebida
                    objItensReqCompra.dQuantPedida = tItemReqCompra.dQuantPedida
                    objItensReqCompra.dQuantRecebida = tItemReqCompra.dQuantRecebida
                    objItensReqCompra.iAlmoxarifado = tItemReqCompra.iAlmoxarifado
                    objItensReqCompra.iCreditaICMS = tItemReqCompra.iCreditaICMS
                    objItensReqCompra.iCreditaIPI = tItemReqCompra.iCreditaIPI
                    objItensReqCompra.iExclusivo = tItemReqCompra.iExclusivo
                    objItensReqCompra.iFilial = tItemReqCompra.iFilial
                    objItensReqCompra.iTipoTributacao = tItemReqCompra.iTipoTributacao
                    objItensReqCompra.lFornecedor = tItemReqCompra.lFornecedor
                    objItensReqCompra.lNumIntDoc = tItemReqCompra.lNumIntDoc
                    objItensReqCompra.lObservacao = tItemReqCompra.lObservacao
                    objItensReqCompra.sCcl = tItemReqCompra.sCcl
                    objItensReqCompra.sContaContabil = tItemReqCompra.sContaContabil
                    objItensReqCompra.sDescProduto = tItemReqCompra.sDescProduto
                    objItensReqCompra.sObservacao = tItemReqCompra.sObservacao
                    objItensReqCompra.sProduto = tItemReqCompra.sProduto
                    objItensReqCompra.sUM = tItemReqCompra.sUM
                    objItensReqCompra.iSelecionado = MARCADO
                    objItensReqCompra.lReqCompra = objReqCompra.lNumIntDoc
                    objItensReqCompra.iItem = objReqCompra.colItens.Count + 1
                    objItensReqCompra.dQuantComprar = objItensReqCompra.dQuantidade - objItensReqCompra.dQuantCancelada - objItensReqCompra.dQuantPedida - objItensReqCompra.dQuantRecebida
            
                    If objItensReqCompra.dQuantComprar > 0.0001 Then objReqCompra.colItens.Add objItensReqCompra
                
                End If
    
                'Busca o Próximo Item
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62737
                
            End If
            
        Loop
    
    Next
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensReqCompra_Le_GeracaoPC = SUCESSO
    
    Exit Function
    
Erro_ItensReqCompra_Le_GeracaoPC:

    ItensReqCompra_Le_GeracaoPC = Err
    
    Select Case Err
    
        Case 62730
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62734, 62735, 62737
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147717)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Private Sub Busca_FilialForn(ByVal colFilialForn As Collection, ByVal lFornecedor As Long, ByVal iFilial As Integer, iPosicao As Integer)

Dim objFilialFornecedor As ClassFilialFornecedor
Dim iIndice As Integer

    iPosicao = 0
    
    For iIndice = 1 To colFilialForn.Count
        
        Set objFilialFornecedor = colFilialForn(iIndice)
        If objFilialFornecedor.lCodFornecedor = lFornecedor And objFilialFornecedor.iCodFilial = iFilial Then
            iPosicao = iIndice
            Exit Sub
        End If
    Next
        
    Exit Sub
    
End Sub

Function Cotacao_Le_NumeroItensRC_Associados(ByVal objCotacao As ClassCotacao, iNumItensRC As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iNumero As Integer

On Error GoTo Erro_Cotacao_Le_NumeroItensRC_Associados

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62810
    
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM CotacaoProdutoItemRC,CotacaoProduto,ItensReqCompra WHERE CotacaoProduto.NumIntDoc = CotacaoProdutoItemRC.CotacaoProduto AND CotacaoProduto.Cotacao = ? AND CotacaoProdutoItemRC.ItemReqCompra = ItensReqCompra.NumIntDoc AND(ItensReqCompra.Quantidade - ItensReqCompra.QuantCancelada - ItensReqCompra.QuantPedida - ItensReqCompra.Quantrecebida) > 0 ", iNumero, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 62811
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 62812
    
    iNumItensRC = iNumero
    
    Call Comando_Fechar(lComando)
    
    Cotacao_Le_NumeroItensRC_Associados = SUCESSO
    
    Exit Function

Erro_Cotacao_Le_NumeroItensRC_Associados:

    Cotacao_Le_NumeroItensRC_Associados = Err
    
    Select Case Err
    
        Case 62810
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62811, 62812
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", Err)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147718)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Cotacao_Le_Codigo(ByVal objCotacao As ClassCotacao) As Long
'Lê a tabela Cotacao utilizando a FilialEmpresa/Codigo como parametro

Dim lErro As Long
Dim lComando As Long
Dim tCotacao As typeCotacao

On Error GoTo Erro_Cotacao_Le_Codigo

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89160

    With tCotacao

        tCotacao.sDescricao = String(STRING_PEDIDO_COTACAO_DESCRICAO, 0)

        'Seleciona os os campos da tabela Cotacao onde objCotacao.lNumIntDoc=objItemPedCotacao.lCotacao
        lErro = Comando_Executar(lComando, "SELECT  NumIntDoc, Descricao, Data, TipoDestino, FornCliDestino, FilialDestino, Comprador FROM Cotacao WHERE FilialEmpresa = ? AND Codigo = ?", tCotacao.lNumIntDoc, tCotacao.sDescricao, tCotacao.dtData, tCotacao.iTipoDestino, tCotacao.lFornCliDestino, tCotacao.iFilialDestino, tCotacao.iComprador, objCotacao.iFilialEmpresa, objCotacao.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 89161

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89162

        If lErro = AD_SQL_SEM_DADOS Then gError 89163

        'Armazena as informações lidas do bd no objeto
        objCotacao.lNumIntDoc = tCotacao.lNumIntDoc
        objCotacao.sDescricao = tCotacao.sDescricao
        objCotacao.dtData = tCotacao.dtData
        objCotacao.iTipoDestino = tCotacao.iTipoDestino
        objCotacao.lFornCliDestino = tCotacao.lFornCliDestino
        objCotacao.iFilialDestino = tCotacao.iFilialDestino
        objCotacao.iComprador = tCotacao.iComprador

    End With

    Call Comando_Fechar(lComando)

    Cotacao_Le_Codigo = SUCESSO

    Exit Function

Erro_Cotacao_Le_Codigo:

    Cotacao_Le_Codigo = gErr

    Select Case gErr

        Case 89160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89161, 89162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case 89163 'Não existe na tabela Cotacao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147719)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompras_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra,cujo numero interno foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim tPedCompra As typePedidoCompras
Dim iEmbalagem As Integer, dTaxa As Double, iMoeda As Integer
On Error GoTo Erro_PedidoCompras_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56115
    'OK  Isso aqui não compila. Essa variável não foi declarada
    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)
'leo
    'Busca na tabela de Pedido de Compra o pedido de Compra com codigo e filial empresa passados
    lErro = Comando_Executar(lComando, "SELECT Embalagem, Taxa, Moeda, NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao, CodigoPV, DataRefFluxo, DataBaixa, TabelaPreco FROM PedidoCompra WHERE NumIntDoc = ?", _
        tPedCompra.iEmbalagem, tPedCompra.dTaxa, tPedCompra.iMoeda, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, tPedCompra.lCodigoPV, tPedCompra.dtDataRefFluxo, tPedCompra.dtDataBaixa, tPedCompra.iTabelaPreco, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 56116

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56117

    'PedidoCompra nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 56118

    'carrega os dados em objPedidoCompra
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.lObservacao = tPedCompra.lObservacao
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.dTaxa = tPedCompra.dTaxa
    objPedidoCompra.iEmbalagem = tPedCompra.iEmbalagem
    objPedidoCompra.iMoeda = tPedCompra.iMoeda
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem
    objPedidoCompra.lCodigoPV = tPedCompra.lCodigoPV
    objPedidoCompra.dtDataRefFluxo = tPedCompra.dtDataRefFluxo
    objPedidoCompra.dtDataBaixa = tPedCompra.dtDataBaixa
    objPedidoCompra.iTabelaPreco = tPedCompra.iTabelaPreco
    
    objPedidoCompra.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PC
    objPedidoCompra.objInfoAdic.lNumIntDoc = objPedidoCompra.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objPedidoCompra.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 56119

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedidoCompras_Le = SUCESSO

    Exit Function

Erro_PedidoCompras_Le:

    PedidoCompras_Le = Err

    Select Case Err

        Case 56115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56116, 56117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra.lCodigo)

        Case 56118
            'erro tratado na rotina chamadora

        Case 56119
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147720)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ConcorrenciaBaixada_Le1(ByVal objConcorrencia As ClassConcorrencia) As Long
'Lê dados da concorrência identificada pelo numintdoc em objConcorrencia
'OBS.: NAO CARREGA OS DADOS DOS ITENS

Dim lErro As Long
Dim lComando As Long
Dim tConcorrencia As typeConcorrencia

On Error GoTo Erro_ConcorrenciaBaixada_Le1

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81621

    tConcorrencia.sDescricao = String(STRING_NOME, 0)

    'Lê dados da concorrência
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Descricao, Data, Comprador, TaxaFinanceira, TipoDestino, FornCliDestino, FilialDestino FROM ConcorrenciaBaixada WHERE NumIntDoc = ?", _
        tConcorrencia.lCodigo, tConcorrencia.iFilialEmpresa, tConcorrencia.sDescricao, tConcorrencia.dtData, tConcorrencia.iComprador, tConcorrencia.dTaxaFinanceira, tConcorrencia.iTipoDestino, tConcorrencia.lFornCliDestino, tConcorrencia.iFilialDestino, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 81622

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81623

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 81624

    objConcorrencia.lCodigo = tConcorrencia.lCodigo
    objConcorrencia.iFilialEmpresa = tConcorrencia.iFilialEmpresa
    objConcorrencia.sDescricao = tConcorrencia.sDescricao
    objConcorrencia.dtData = tConcorrencia.dtData
    objConcorrencia.iComprador = tConcorrencia.iComprador
    objConcorrencia.dTaxaFinanceira = tConcorrencia.dTaxaFinanceira
    objConcorrencia.iTipoDestino = tConcorrencia.iTipoDestino
    objConcorrencia.lFornCliDestino = tConcorrencia.lFornCliDestino
    objConcorrencia.iFilialDestino = tConcorrencia.iFilialDestino

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ConcorrenciaBaixada_Le1 = SUCESSO

    Exit Function

Erro_ConcorrenciaBaixada_Le1:

    ConcorrenciaBaixada_Le1 = gErr

    Select Case gErr

        Case 81621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81622, 81623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case 81624 'Concorrência não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147721)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CotacaoItemConcorrenciaTodos_Le(ByVal objItemConcorrencia As ClassItemConcorrencia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tCotacaoItemConc As typeCotacaoItemConcorrencia
Dim objCotacaoItemConc As ClassCotacaoItemConc
Dim sSelect As String
Dim iFilial As Integer
Dim iCondPagto As Integer
Dim dQuantCotada As Double
Dim dtDataCotacao As Date

On Error GoTo Erro_CotacaoItemConcorrenciaTodos_Le

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86155

    sSelect = "SELECT CotacaoItemConcorrenciaTodos.ItemCotacao,CotacaoItemConcorrenciaTodos.NumIntDoc,CotacaoItemConcorrenciaTodos.ValorPresente,CotacaoItemConcorrenciaTodos.Escolhido,CotacaoItemConcorrenciaTodos.MotivoEscolha,CotacaoItemConcorrenciaTodos.QuantidadeComprar,CotacaoItemConcorrenciaTodos.DataEntrega,CotacaoItemConcorrenciaTodos.PrecoAjustado,ItensCotacaoTodos.CondPagto, ItensCotacaoTodos.AliquotaIPI,ItensCotacaoTodos.AliquotaICMS, Fornecedores.NomeReduzido,FiliaisFornecedores.CodFilial, FiliaisFornecedores.Nome,PedidoCotacaoTodos.Codigo, CotacaoProdutoTodas.Quantidade, CondicoesPagto.DescReduzida,PedidoCotacaoTodos.DataValidade,CotacaoTodas.Data,ItensCotacaoTodos.PrazoEntrega,ItensCotacaoTodos.QuantEntrega,ItensCotacaoTodos.PrecoUnitario, ItensCotacaoTodos.Moeda, ItensCotacaoTodos.Taxa " & _
              " FROM CotacaoItemConcorrenciaTodos, ItensCotacaoTodos, ItemPedCotacaoTodos, PedidoCotacaoTodos, Fornecedores, FiliaisFornecedores, CotacaoProdutoTodas, CondicoesPagto, CotacaoTodas " & _
              " WHERE ItensCotacaoTodos.NumIntDoc = CotacaoItemConcorrenciaTodos.ItemCotacao AND ItemPedCotacaoTodos.NumIntDoc = ItensCotacaoTodos.ItemPedCotacao AND PedidoCotacaoTodos.NumIntDoc = ItemPedCotacaoTodos.PedidoCotacao AND PedidoCotacaoTodos.Fornecedor = fornecedores.codigo AND FiliaisFornecedores.CodFilial = PedidoCotacaoTodos.Filial AND FiliaisFornecedores.codfornecedor = Fornecedores.Codigo AND CotacaoprodutoTodas.NumIntDoc = ItemPedCotacaoTodos.CotacaoProduto AND CondicoesPagto.Codigo = ItensCotacaoTodos.CondPagto AND CotacaoTodas.numIntDoc = CotacaoProdutoTodas.Cotacao AND CotacaoItemConcorrenciaTodos.Itemconcorrencia = ?"
        
    tCotacaoItemConc.sMotivoEscolha = String(STRING_MOTIVO_ESCOLHA, 0)
    tCotacaoItemConc.sFilial = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCotacaoItemConc.sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    tCotacaoItemConc.sCondPagto = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
    

    'Busca em CotacaoItemConcorrencia os registros ligados ao item passado
    lErro = Comando_Executar(lComando, sSelect, tCotacaoItemConc.lItemCotacao, tCotacaoItemConc.lNumIntDoc, tCotacaoItemConc.dValorPresente, tCotacaoItemConc.iEscolhido, tCotacaoItemConc.sMotivoEscolha, tCotacaoItemConc.dQuantidadeComprar, tCotacaoItemConc.dtDataEntrega, tCotacaoItemConc.dPrecoAjustado, iCondPagto, tCotacaoItemConc.dAliquotaIPI, tCotacaoItemConc.dCreditoICMS, tCotacaoItemConc.sFornecedor, iFilial, tCotacaoItemConc.sFilial, tCotacaoItemConc.lPedCotacao, dQuantCotada, tCotacaoItemConc.sCondPagto, tCotacaoItemConc.dtDataValidade, dtDataCotacao, tCotacaoItemConc.iPrazoEntrega, tCotacaoItemConc.dQuantEntrega, tCotacaoItemConc.dPrecoUnitario, tCotacaoItemConc.iMoeda, tCotacaoItemConc.dTaxa, objItemConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 86156

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86157

    'Para cada registro encontrado
    Do While lErro <> AD_SQL_SEM_DADOS

        'Preenche um objCotacaoItemConc com os dados lidos
        Set objCotacaoItemConc = New ClassCotacaoItemConc

        With tCotacaoItemConc

            objCotacaoItemConc.dtDataValidade = .dtDataValidade
            objCotacaoItemConc.dtDataPedidoCotacao = dtDataCotacao
            objCotacaoItemConc.dCreditoICMS = .dCreditoICMS
            objCotacaoItemConc.dCreditoIPI = .dCreditoIPI
            objCotacaoItemConc.dPrecoAjustado = .dPrecoAjustado
            objCotacaoItemConc.dPrecoUnitario = .dPrecoUnitario
            objCotacaoItemConc.dPreferencia = .dPreferencia
            objCotacaoItemConc.dQuantEntrega = .dQuantEntrega
            objCotacaoItemConc.dQuantidadeComprar = .dQuantidadeComprar
            objCotacaoItemConc.dtDataEntrega = .dtDataEntrega
            objCotacaoItemConc.dtDataValidade = .dtDataValidade
            objCotacaoItemConc.dValorPresente = .dValorPresente
            objCotacaoItemConc.iEscolhido = .iEscolhido
            objCotacaoItemConc.iSelecionada = MARCADO
            objCotacaoItemConc.iPrazoEntrega = .iPrazoEntrega
            objCotacaoItemConc.lNumIntDoc = .lNumIntDoc
            objCotacaoItemConc.lItemCotacao = .lItemCotacao
            objCotacaoItemConc.sMotivoEscolha = .sMotivoEscolha
            objCotacaoItemConc.dAliquotaICMS = .dAliquotaICMS
            objCotacaoItemConc.dAliquotaIPI = .dAliquotaIPI
            objCotacaoItemConc.sCondPagto = iCondPagto & SEPARADOR & tCotacaoItemConc.sCondPagto '??? ler s
            objCotacaoItemConc.sFilial = iFilial & SEPARADOR & tCotacaoItemConc.sFilial
            objCotacaoItemConc.sFornecedor = .sFornecedor
            objCotacaoItemConc.lPedCotacao = .lPedCotacao
            objCotacaoItemConc.dQuantCotada = dQuantCotada
            objCotacaoItemConc.iMoeda = .iMoeda
            objCotacaoItemConc.dTaxa = .dTaxa
            
        End With

        'Adiciona o objCotacaoItemConc na coleção do item de concorrência
        objItemConcorrencia.colCotacaoItemConc.Add objCotacaoItemConc

        'Busca o proximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86158

    Loop

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    CotacaoItemConcorrenciaTodos_Le = SUCESSO

    Exit Function

Erro_CotacaoItemConcorrenciaTodos_Le:

    CotacaoItemConcorrenciaTodos_Le = gErr

    Select Case gErr

        Case 86155
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86156, 86157, 86158
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147722)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UltimaCotacao_Le_FornecedorProduto(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados da ultima cotacao (maior data de pedido de cotação) do produto/Fornecedor/FilialForn feita pela FilialEmpresa passado como parametro.

Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF

On Error GoTo Erro_UltimaCotacao_Le_FornecedorProduto

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89326

    sSelect = "SELECT ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.DataEmissao, PedidoCotacao.TipoFrete, CotacaoProduto.Quantidade, CotacaoProduto.UM, CondicoesPagto.DescReduzida FROM PedidoCotacao, ItemPedCotacao, CotacaoProduto, ItensCotacao, CondicoesPagto Where PedidoCotacao.NumIntDoc = ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto = CotacaoProduto.NumIntDoc AND ItemPedCotacao.NumIntDoc = ItensCotacao.ItemPedCotacao AND CondicoesPagto.Codigo = ItensCotacao.CondPagto AND PedidoCotacao.Fornecedor = ? And PedidoCotacao.Filial = ? AND CotacaoProduto.Produto = ? AND PedidoCotacao.FilialEmpresa = ? ORDER BY PedidoCotacao.DataEmissao DESC"
        
    With tFornecedorProdutoFF

        .sCondPagto = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
        .sUMQuantUltimaCotacao = String(STRING_UM_SIGLA, 0)

        'Busca em CotacaoItemConcorrencia os registros ligados ao item passado
        lErro = Comando_Executar(lComando, sSelect, .iCondPagto, .dUltimaCotacao, .dtDataUltimaCotacao, .iTipoFreteUltimaCotacao, .dQuantUltimaCotacao, .sUMQuantUltimaCotacao, .sCondPagto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 89327

    End With

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89328

    If lErro = AD_SQL_SUCESSO Then

        With tFornecedorProdutoFF
    
            objFornecedorProdutoFF.iCondPagto = .iCondPagto
            objFornecedorProdutoFF.dUltimaCotacao = .dUltimaCotacao
            objFornecedorProdutoFF.dtDataUltimaCotacao = .dtDataUltimaCotacao
            objFornecedorProdutoFF.iTipoFreteUltimaCotacao = .iTipoFreteUltimaCotacao
            objFornecedorProdutoFF.dQuantUltimaCotacao = .dQuantUltimaCotacao
            objFornecedorProdutoFF.sCondPagto = .sCondPagto
            objFornecedorProdutoFF.sUMQuantUltimaCotacao = .sUMQuantUltimaCotacao
            
        End With
    
    ElseIf lErro = AD_SQL_SEM_DADOS Then
    
        objFornecedorProdutoFF.dtDataUltimaCotacao = DATA_NULA

    End If

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    UltimaCotacao_Le_FornecedorProduto = SUCESSO

    Exit Function

Erro_UltimaCotacao_Le_FornecedorProduto:

    UltimaCotacao_Le_FornecedorProduto = gErr

    Select Case gErr

        Case 89326
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89327, 89328
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULTIMA_COTACAO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147723)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function UltimaCompra_Le_FornecedorProduto(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados da ultima compra global para o produto/Fornecedor/FilialForn passado como parametro.

Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim dValorDesconto As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double

On Error GoTo Erro_UltimaCompra_Le_FornecedorProduto

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89347

    tFornecedorProdutoFF.sUMQuantPedida = String(STRING_UM_SIGLA, 0)
        '
    With tFornecedorProdutoFF
'
        'Alterado por Wagner
        'Busca em NFiscal, ItensNFiscal os registros ligados ao item passado
        lErro = Comando_Executar(lComando, "SELECT ItensNFiscal.PrecoUnitario, ItensNFiscal.Quantidade, ItensNFiscal.ValorDesconto, PedidoCompraN.DataEnvio, NFiscal.DataEntrada, ItensPedCompraN.Quantidade, ItensPedCompraN.UM FROM ItensNFiscal, NFiscal, PedidoCompraN, ItemNFItemPC, ItensPedCompraN WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND PedidoCompraN.NumIntDoc = ItensPedCompraN.PedCompra AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompraN.NumIntDoc = ItemNFItemPC.ItemPedCompra AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND ItensNFiscal.Produto = ? AND NFiscal.Status <> ? AND NFiscal.Status <> ? AND PedidoCompraN.FilialEmpresa = ? ORDER BY PedidoCompraN.DataEnvio DESC", dPrecoUnitario, dQuantidade, dValorDesconto, .dtDataUltimaCompra, .dtDataReceb, .dQuantPedida, .sUMQuantPedida, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, STATUS_EXCLUIDO, STATUS_CANCELADO, objFornecedorProdutoFF.iFilialEmpresa) 'Wagner
        If lErro <> AD_SQL_SUCESSO Then gError 89348

    End With

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89349
    If lErro = AD_SQL_SUCESSO Then

        With tFornecedorProdutoFF
            
            objFornecedorProdutoFF.dtDataReceb = .dtDataReceb
            objFornecedorProdutoFF.dtDataUltimaCompra = .dtDataUltimaCompra
            objFornecedorProdutoFF.dPrecoTotal = (dPrecoUnitario * dQuantidade) - dValorDesconto
            objFornecedorProdutoFF.dQuantPedida = .dQuantPedida
            objFornecedorProdutoFF.dQuantRecebida = dQuantidade
            objFornecedorProdutoFF.sUMQuantPedida = .sUMQuantPedida
            
        End With

    Else
        objFornecedorProdutoFF.dtDataReceb = DATA_NULA
        objFornecedorProdutoFF.dtDataUltimaCompra = DATA_NULA
        objFornecedorProdutoFF.dPrecoTotal = 0
        objFornecedorProdutoFF.dQuantPedida = 0
        objFornecedorProdutoFF.dQuantRecebida = 0
        objFornecedorProdutoFF.sUMQuantPedida = ""
        
    End If

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    UltimaCompra_Le_FornecedorProduto = SUCESSO

    Exit Function

Erro_UltimaCompra_Le_FornecedorProduto:

    UltimaCompra_Le_FornecedorProduto = gErr

    Select Case gErr

        Case 89347
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89348, 89349
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULTIMA_COMPRA", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147724)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function UltimoItemPedCompraFechado_Le(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados do ultimo (mais recente nota fiscal) item de pedido de compra da Filial em questão com status fechado para o produto/Fornecedor/FilialForn passado como parametro.


Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim dValorDesconto As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double

On Error GoTo Erro_UltimoItemPedCompraFechado_Le

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89350

    sSelect = "SELECT NFiscal.DataEntrada, PedidoCompraTodos.Data, ItensPedCompraTodos.QuantRecebida, ItensPedCompraTodos.Quantidade, ItensPedCompraTodos.UM FROM PedidoCompraTodos, ItensPedCompraTodos, ItensNFiscal, NFiscal, ItemNFItemPC  WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItemNFItemPC.ItemPedCompra = ItensPedCompraTodos.NumIntDoc AND ItensPedCompraTodos.PedCompra = PedidoCompraTodos.NumIntDoc AND ItensPedCompraTodos.Status = 1 AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND ItensNFiscal.Produto = ? AND PedidoCompraTodos.FilialEmpresa = ? ORDER BY NFiscal.DataEntrada DESC"
        
    With tFornecedorProdutoFF

        .sUMQuantPedida = String(STRING_UM_SIGLA, 0)
        .sUMQuantRecebida = String(STRING_UM_SIGLA, 0)

        'Busca em NFiscal, ItensNFiscal os registros ligados ao item passado
        lErro = Comando_Executar(lComando, sSelect, .dtDataReceb, .dtDataPedido, .dQuantRecebida, .dQuantPedida, .sUMQuantPedida, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 89351

    End With

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89352

    If lErro = AD_SQL_SUCESSO Then

        With tFornecedorProdutoFF
    
            objFornecedorProdutoFF.dQuantPedida = .dQuantPedida
            objFornecedorProdutoFF.dQuantRecebida = .dQuantRecebida
            objFornecedorProdutoFF.dtDataPedido = .dtDataPedido
            objFornecedorProdutoFF.dtDataReceb = .dtDataReceb
            objFornecedorProdutoFF.sUMQuantPedida = .sUMQuantPedida
            objFornecedorProdutoFF.sUMQuantRecebida = .sUMQuantPedida
            
        End With
    
    ElseIf lErro = AD_SQL_SEM_DADOS Then
        
        objFornecedorProdutoFF.dQuantPedida = 0
        objFornecedorProdutoFF.dQuantRecebida = 0
        objFornecedorProdutoFF.dtDataPedido = DATA_NULA
        objFornecedorProdutoFF.dtDataReceb = DATA_NULA
        objFornecedorProdutoFF.sUMQuantPedida = ""
        objFornecedorProdutoFF.sUMQuantRecebida = ""
        
    End If

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    UltimoItemPedCompraFechado_Le = SUCESSO

    Exit Function

Erro_UltimoItemPedCompraFechado_Le:

    UltimoItemPedCompraFechado_Le = gErr

    Select Case gErr

        Case 89350
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89351, 89352
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULTIMO_PEDCOMPRA_FECHADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147725)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function QuantProdutoPedAbertos_Le(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna a Quantidade em Unidade de Compra do Produto Pedido pela Filial em questão e ainda não entregue para o produto/Fornecedor/FilialForn passado como parametro.

Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim dQuantidadePedida As Double
Dim dQuantidadeRecebida As Double
Dim objProduto As New ClassProduto
Dim sUM As String
Dim dFator As Double

On Error GoTo Erro_QuantProdutoPedAbertos_Le

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89360

    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    'Le no Produto a Unidade de Compras
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 89361
    
    If lErro = 28030 Then gError 89362

    sSelect = "SELECT ItensPedCompra.Quantidade, ItensPedCompra.QuantRecebida, ItensPedCompra.UM FROM PedidoCompra, ItensPedCompra  WHERE ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND ItensPedCompra.Status = 0 AND PedidoCompra.Fornecedor = ? AND PedidoCompra.Filial = ? AND ItensPedCompra.Produto = ? AND PedidoCompra.FilialEmpresa = ?"
        
    sUM = String(STRING_UM_SIGLA, 0)
        
    With tFornecedorProdutoFF

        'Busca em NFiscal, ItensNFiscal os registros ligados ao item passado
        lErro = Comando_Executar(lComando, sSelect, dQuantidadePedida, dQuantidadeRecebida, sUM, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 89363

    End With

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89364

    Do While lErro = AD_SQL_SUCESSO

        With tFornecedorProdutoFF
    
            'Converte as unidades de medida
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMCompra, dFator)
            If lErro <> SUCESSO Then gError 89366
    
            .dQuantPedAbertos = .dQuantPedAbertos + ((dQuantidadePedida - dQuantidadeRecebida) * dFator)
    
        End With
    
        'Busca o proximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89365

    Loop

    objFornecedorProdutoFF.dQuantPedAbertos = tFornecedorProdutoFF.dQuantPedAbertos

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    QuantProdutoPedAbertos_Le = SUCESSO

    Exit Function

Erro_QuantProdutoPedAbertos_Le:

    QuantProdutoPedAbertos_Le = gErr

    Select Case gErr

        Case 89360
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89361, 89366

        Case 89362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 89363, 89364, 89365
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_PEDCOMPRA_ABERTO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147726)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TempoRessup_Calcula(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Calcula o Tempo de Ressuprimento e devolve em objFornecedorProdutoFF

Dim lErro As Long
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim alComando(2) As Long
Dim dQuantidade As Double
Dim dtDataEntrada As Date
Dim dtDataEnvio As Date
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dTempoEntregaTotal As Double
Dim dQuantEntregaTotal As Long
Dim iIndice As Integer
Dim sSQL As String
Dim iNotasFiscaisPrimeiras As Integer
Dim sUMItemPC As String

On Error GoTo Erro_TempoRessup_Calcula

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67039
    Next
    
    'Lê o Numero de Compras de TempoRessup
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 67040
    
    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    'Le no Produto a Unidade de Compras
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 67043
    
    If lErro = 28030 Then gError 67044
    
    sSQL = "SELECT ItemNFItemPC.Quantidade, NFiscal.DataEntrada, PedidoCompraTodos.DataEnvio, ItensPedCompraTodos.UM FROM NFiscal, ItensNFiscal, ItemNFItemPC, PedidoCompraTodos, ItensPedCompraTodos WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItemNFItemPC.ItemPedCompra = ItensPedCompraTodos.NumIntDoc AND PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND ItensNFiscal.Produto = ? AND PedidoCompraTodos.DataEnvio <> ? AND (TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ?) ORDER BY NFiscal.DataEntrada DESC"
    
    sUMItemPC = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(1), sSQL, dQuantidade, dtDataEntrada, dtDataEnvio, sUMItemPC, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, DATA_NULA, DOCINFO_NFIENCO, DOCINFO_NFEENCO, DOCINFO_NFIEFNCO, DOCINFO_NFEEFNCO)
    If lErro <> AD_SQL_SUCESSO Then gError 67041
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67042
    
    iNotasFiscaisPrimeiras = 0
    
    Do While (lErro = SUCESSO And iNotasFiscaisPrimeiras < objConfiguraCOM.iNumComprasTempoRessup)
                
        iNotasFiscaisPrimeiras = iNotasFiscaisPrimeiras + 1
        
        'Converte as unidades de medida
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItemPC, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 67045
                
        dQuantidade = dQuantidade * dFator
        
        dTempoEntregaTotal = dTempoEntregaTotal + (dQuantidade * (dtDataEntrada - dtDataEnvio))
        
        dQuantEntregaTotal = dQuantEntregaTotal + dQuantidade
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67046
    
    Loop
    
    If dQuantEntregaTotal = 0 Then
        objFornecedorProdutoFF.dTempoRessup = 0
    Else
        objFornecedorProdutoFF.dTempoRessup = dTempoEntregaTotal / dQuantEntregaTotal
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TempoRessup_Calcula = SUCESSO
    
    Exit Function
    
Erro_TempoRessup_Calcula:

    TempoRessup_Calcula = gErr
    
    Select Case gErr
        
        Case 67039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 67040, 67043, 67045
        
        Case 67041, 67042, 67046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
        
        Case 67044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147727)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornProd(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/Produto em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo ERRO_PESQUISA_PRIMEIRA_COMPRA_FORNPROD

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89371
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89372
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89373
    
    If lErro = AD_SQL_SUCESSO Then
        objFornecedorProdutoFF.dtDataPrimeiraCompra = dtDataEmissao
    Else
        objFornecedorProdutoFF.dtDataPrimeiraCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Primeira_Compra_FornProd = SUCESSO
    
    Exit Function
    
ERRO_PESQUISA_PRIMEIRA_COMPRA_FORNPROD:

    Pesquisa_Primeira_Compra_FornProd = gErr
    
    Select Case gErr
        
        Case 89371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89372, 89373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_PRIM_COMPRA_FORNPROD", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147728)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornProd(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/Produto em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo ERRO_PESQUISA_ULTIMA_COMPRA_FORNPROD

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89374
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89375
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89376
    
    If lErro = AD_SQL_SUCESSO Then
        objFornecedorProdutoFF.dtDataUltimaCompra = dtDataEmissao
    Else
        objFornecedorProdutoFF.dtDataUltimaCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Compra_FornProd = SUCESSO
    
    Exit Function
    
ERRO_PESQUISA_ULTIMA_COMPRA_FORNPROD:

    Pesquisa_Ultima_Compra_FornProd = gErr
    
    Select Case gErr
        
        Case 89374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89375, 89376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_COMPRA_FORNPROD", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147729)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornFilProd(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/FilialForn/Produto em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Primeira_Compra_FornFilProd

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89377
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialForn = ? AND NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89378
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89379
    
    If lErro = AD_SQL_SUCESSO Then
        objFornecedorProdutoFF.dtDataPrimeiraCompra = dtDataEmissao
    Else
        objFornecedorProdutoFF.dtDataPrimeiraCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Primeira_Compra_FornFilProd = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Primeira_Compra_FornFilProd:

    Pesquisa_Primeira_Compra_FornFilProd = gErr
    
    Select Case gErr
        
        Case 89376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89378, 89379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_PRIM_COMPRA_FORNFILPROD", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147730)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornFilProd(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn/Produto em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Ultima_Compra_FornFilProd

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89380
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialForn = ? AND NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89381
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89382
    
    If lErro = AD_SQL_SUCESSO Then
        objFornecedorProdutoFF.dtDataUltimaCompra = dtDataEmissao
    Else
        objFornecedorProdutoFF.dtDataUltimaCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Compra_FornFilProd = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Ultima_Compra_FornFilProd:

    Pesquisa_Ultima_Compra_FornFilProd = gErr
    
    Select Case gErr
        
        Case 89380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89381, 89382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_COMPRA_FORNFILPROD", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147731)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_Forn(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Primeira_Compra_Forn

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89386
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89387
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89388
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataPrimeiraCompra = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataPrimeiraCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Primeira_Compra_Forn = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Primeira_Compra_Forn:

    Pesquisa_Primeira_Compra_Forn = gErr
    
    Select Case gErr
        
        Case 89386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89387, 89388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_PRIM_COMPRA_FORN", gErr, objFilialFornecedorEst.lCodFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147732)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_Forn(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Ultima_Compra_Forn

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89389
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89390
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89391
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataUltimaCompra = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataUltimaCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Compra_Forn = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Ultima_Compra_Forn:

    Pesquisa_Ultima_Compra_Forn = gErr
    
    Select Case gErr
        
        Case 89389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89390, 89391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_COMPRA_FORN", gErr, objFilialFornecedorEst.lCodFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147733)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornFil(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/FilialForn em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Primeira_Compra_FornFil

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89392
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialForn = ? AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89393
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89394
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataPrimeiraCompra = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataPrimeiraCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Primeira_Compra_FornFil = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Primeira_Compra_FornFil:

    Pesquisa_Primeira_Compra_FornFil = gErr
    
    Select Case gErr
        
        Case 89392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89393, 89394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_PRIM_COMPRA_FORNFIL", gErr, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147734)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornFil(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Ultima_Compra_FornFil

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89395
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Compras = 1 AND NFiscal.Fornecedor = ? And NFiscal.FilialForn = ? AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89396
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89397
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataUltimaCompra = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataUltimaCompra = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Compra_FornFil = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Ultima_Compra_FornFil:

    Pesquisa_Ultima_Compra_FornFil = gErr
    
    Select Case gErr
        
        Case 89395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89396, 89397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_COMPRA_FORNFIL", gErr, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147735)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TempoRessupFilial_Calcula(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Calcula o Tempo de Ressuprimento da Filial e devolve em objFornecedorProdutoFF

Dim lErro As Long
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim alComando(2) As Long
Dim dQuantidade As Double
Dim dtDataEntrada As Date
Dim dtDataEnvio As Date
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dTempoEntregaTotal As Double
Dim dQuantEntregaTotal As Long
Dim iIndice As Integer
Dim sSQL As String
Dim iNotasFiscaisPrimeiras As Integer
Dim sUMItemPC As String

On Error GoTo Erro_TempoRessupFilial_Calcula

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67039
    Next
    
    'Lê o Numero de Compras de TempoRessup
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 67040
    
    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    'Le no Produto a Unidade de Compras
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 67043
    
    If lErro = 28030 Then gError 67044
    
    sSQL = "SELECT ItemNFItemPC.Quantidade, NFiscal.DataEntrada, PedidoCompraTodos.DataEnvio, ItensPedCompraTodos.UM FROM NFiscal, ItensNFiscal, ItemNFItemPC, PedidoCompraTodos, ItensPedCompraTodos WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItemNFItemPC.ItemPedCompra = ItensPedCompraTodos.NumIntDoc AND PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND NFiscal.FilialEmpresa = ? AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND ItensNFiscal.Produto = ? AND PedidoCompraTodos.DataEnvio <> ? AND (TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ?) ORDER BY NFiscal.DataEntrada DESC"
    
    sUMItemPC = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(1), sSQL, dQuantidade, dtDataEntrada, dtDataEnvio, sUMItemPC, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, DATA_NULA, DOCINFO_NFIENCO, DOCINFO_NFEENCO, DOCINFO_NFIEFNCO, DOCINFO_NFEEFNCO)
    If lErro <> AD_SQL_SUCESSO Then gError 67041
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67042
    
    iNotasFiscaisPrimeiras = 0
    
    Do While (lErro = SUCESSO And iNotasFiscaisPrimeiras < objConfiguraCOM.iNumComprasTempoRessup)
                
        iNotasFiscaisPrimeiras = iNotasFiscaisPrimeiras + 1
        
        'Converte as unidades de medida
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItemPC, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 67045
                
        dQuantidade = dQuantidade * dFator
        
        dTempoEntregaTotal = dTempoEntregaTotal + (dQuantidade * (dtDataEntrada - dtDataEnvio))
        
        dQuantEntregaTotal = dQuantEntregaTotal + dQuantidade
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67046
    
    Loop
    
    If dQuantEntregaTotal = 0 Then
        objFornecedorProdutoFF.dTempoRessup = 0
    Else
        objFornecedorProdutoFF.dTempoRessup = dTempoEntregaTotal / dQuantEntregaTotal
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TempoRessupFilial_Calcula = SUCESSO
    
    Exit Function
    
Erro_TempoRessupFilial_Calcula:

    TempoRessupFilial_Calcula = gErr
    
    Select Case gErr
        
        Case 67039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 67040, 67043, 67045
        
        Case 67041, 67042, 67046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
        
        Case 67044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147736)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ConcorrenciaN_Le(ByVal objConcorrencia As ClassConcorrencia) As Long
'Lê dados da concorrência passada pelo código e filialempresa em objConcorrencia

Dim lErro As Long
Dim lComando As Long
Dim tConcorrencia As typeConcorrencia

On Error GoTo Erro_ConcorrenciaN_Le

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89862

    tConcorrencia.sDescricao = String(STRING_NOME, 0)

    'Lê dados da concorrência
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Descricao, Data, Comprador, TaxaFinanceira, TipoDestino, FornCliDestino, FilialDestino FROM ConcorrenciaN WHERE Codigo  = ? AND FilialEmpresa = ?", _
    tConcorrencia.lNumIntDoc, tConcorrencia.iFilialEmpresa, tConcorrencia.sDescricao, tConcorrencia.dtData, tConcorrencia.iComprador, tConcorrencia.dTaxaFinanceira, tConcorrencia.iTipoDestino, tConcorrencia.lFornCliDestino, tConcorrencia.iFilialDestino, objConcorrencia.lCodigo, objConcorrencia.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 89863

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89864

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 89865

    objConcorrencia.lNumIntDoc = tConcorrencia.lNumIntDoc
    objConcorrencia.iFilialEmpresa = tConcorrencia.iFilialEmpresa
    objConcorrencia.sDescricao = tConcorrencia.sDescricao
    objConcorrencia.dtData = tConcorrencia.dtData
    objConcorrencia.iComprador = tConcorrencia.iComprador
    objConcorrencia.dTaxaFinanceira = tConcorrencia.dTaxaFinanceira
    objConcorrencia.iTipoDestino = tConcorrencia.iTipoDestino
    objConcorrencia.lFornCliDestino = tConcorrencia.lFornCliDestino
    objConcorrencia.iFilialDestino = tConcorrencia.iFilialDestino

    'Lê os itens da Concorrência
    lErro = CF("ItensConcorrenciaTodos_Le", objConcorrencia)
    If lErro <> SUCESSO Then gError 89866

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ConcorrenciaN_Le = SUCESSO

    Exit Function

Erro_ConcorrenciaN_Le:

    ConcorrenciaN_Le = gErr

    Select Case gErr

        Case 89862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89863, 89864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case 89865, 89866

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147737)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedCompraGerado_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra Gerado,cujo numero interno foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim tPedCompra As typePedidoCompras

On Error GoTo Erro_PedCompraGerado_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89944
    
    'OK  Isso aqui não compila. Essa variável não foi declarada
    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    tPedCompra.sObsEmbalagem = String(STRING_OBSEMBALAGEM, 0)

    'Busca na tabela de Pedido de Compra o pedido de Compra com codigo e filial empresa passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao FROM PedCompraGerado WHERE NumIntDoc = ?", tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
    tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.sObsEmbalagem, tPedCompra.lObservacao, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89945

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89946

    'PedidoCompra nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 89947

    'carrega os dados em objPedidoCompra
    objPedidoCompra.iFilialEmpresa = tPedCompra.iFilialEmpresa
    objPedidoCompra.lCodigo = tPedCompra.lCodigo
    objPedidoCompra.lFornecedor = tPedCompra.lFornecedor
    objPedidoCompra.iFilial = tPedCompra.iFilial
    objPedidoCompra.iComprador = tPedCompra.iComprador
    objPedidoCompra.sContato = tPedCompra.sContato
    objPedidoCompra.iTipoDestino = tPedCompra.iTipoDestino
    objPedidoCompra.lFornCliDestino = tPedCompra.lFornCliDestino
    objPedidoCompra.iFilialDestino = tPedCompra.iFilialDestino
    objPedidoCompra.dtData = tPedCompra.dtData
    objPedidoCompra.dtDataEmissao = tPedCompra.dtDataEmissao
    objPedidoCompra.dtDataEnvio = tPedCompra.dtDataEnvio
    objPedidoCompra.dtDataAlteracao = tPedCompra.dtDataAlteracao
    objPedidoCompra.iCondicaoPagto = tPedCompra.iCondicaoPagto
    objPedidoCompra.dOutrasDespesas = tPedCompra.dOutrasDespesas
    objPedidoCompra.dValorFrete = tPedCompra.dValorFrete
    objPedidoCompra.dValorSeguro = tPedCompra.dValorSeguro
    objPedidoCompra.dValorDesconto = tPedCompra.dValorDesconto
    objPedidoCompra.dValorTotal = tPedCompra.dValorTotal
    objPedidoCompra.dValorIPI = tPedCompra.dValorIPI
    objPedidoCompra.lObservacao = tPedCompra.lObservacao
    objPedidoCompra.sTipoFrete = tPedCompra.sTipoFrete
    objPedidoCompra.iTransportadora = tPedCompra.iTransportadora
    objPedidoCompra.iProxSeqBloqueio = tPedCompra.iProxSeqBloqueio
    objPedidoCompra.iTipoBaixa = tPedCompra.iTipoBaixa
    objPedidoCompra.sMotivoBaixa = tPedCompra.sMotivoBaixa
    objPedidoCompra.sAlcada = tPedCompra.sAlcada
    objPedidoCompra.sObsEmbalagem = tPedCompra.sObsEmbalagem

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PedCompraGerado_Le = SUCESSO

    Exit Function

Erro_PedCompraGerado_Le:

    PedCompraGerado_Le = gErr

    Select Case gErr

        Case 89944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89945, 89946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDCOMPRAGERADO", gErr, objPedidoCompra.lNumIntDoc)

        Case 89947

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147738)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Janaina - 18/07 - Estava em GeracaoPedCotacaoOcx
Function ItemReqCompras_Le_Primeiro(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê o primeiro Ítem da Requisição Compras passada por objRequisiçãoCompras
'Retorna na colecao de Itens

Dim lErro As Long
Dim tItemReqCompras As typeItemReqCompra
Dim objItemReqCompras As New ClassItemReqCompras
Dim lComando As Long

On Error GoTo Erro_ItemReqCompras_Le_Primeiro

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25972

    With tItemReqCompras

        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabil = String(STRING_CONTA, 0)
    
    End With

   'Lê o ítem da Requisição na tabela ItensReqCompras
    With tItemReqCompras
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompra WHERE ReqCompra = ? ORDER BY NumIntDoc", _
            .lNumIntDoc, .lReqCompra, .sProduto, .sDescProduto, .iStatus, .dQuantidade, .dQuantPedida, .dQuantRecebida, .dQuantCancelada, .sUM, .sCcl, .iAlmoxarifado, .sContaContabil, .iTipoTributacao, .lObservacao, .lFornecedor, .iFilial, .iExclusivo, objRequisicaoCompras.lNumIntDoc)
    End With
        
    If lErro <> AD_SQL_SUCESSO Then Error 25973
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25974

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 25975

    Set objItemReqCompras = New ClassItemReqCompras

    With tItemReqCompras
        
        'Preenche o objItemReqCompras
        objItemReqCompras.lNumIntDoc = .lNumIntDoc
        objItemReqCompras.lReqCompra = .lReqCompra
        objItemReqCompras.sProduto = .sProduto
        objItemReqCompras.sDescProduto = .sDescProduto
        objItemReqCompras.iStatus = .iStatus
        objItemReqCompras.dQuantidade = .dQuantidade
        objItemReqCompras.dQuantPedida = .dQuantPedida
        objItemReqCompras.dQuantRecebida = .dQuantRecebida
        objItemReqCompras.dQuantCancelada = .dQuantCancelada
        objItemReqCompras.sUM = .sUM
        objItemReqCompras.sCcl = .sCcl
        objItemReqCompras.iAlmoxarifado = .iAlmoxarifado
        objItemReqCompras.sContaContabil = .sContaContabil
        objItemReqCompras.iTipoTributacao = .iTipoTributacao
        objItemReqCompras.lObservacao = .lObservacao
        objItemReqCompras.lFornecedor = .lFornecedor
        objItemReqCompras.iFilial = .iFilial
        objItemReqCompras.iExclusivo = .iExclusivo
    
    End With

    'Adiciona o objItemReqCompras na coleção colItens do objRequisiçãoCompras
    objRequisicaoCompras.colItens.Add objItemReqCompras

    Call Comando_Fechar(lComando)

    ItemReqCompras_Le_Primeiro = SUCESSO

    Exit Function

Erro_ItemReqCompras_Le_Primeiro:

    ItemReqCompras_Le_Primeiro = Err

    Select Case Err

        Case 25972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25973, 25974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 25975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRAS_AUSENCIA_ITENS", Err, objRequisicaoCompras.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147739)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Janaina - 18/07 - Estava em GeracaoPedCotacaoOcx
Function ItemRC_NumInt_Le(ByVal objItemReqCompra As ClassItemReqCompras, ByVal iFilialReqCompra As Integer) As Long
'Lê Numintdoc do Item da Requisição de Compras a partir do código da Requisição, FilialEmpresa e
'Produto, Fornecedor, Filial e exclusividade do item da requisição

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemRC_NumInt_Le
    
    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 70395
    
    'Procura no BD o Item da Requisição passado
    lErro = Comando_Executar(lComando, "SELECT ItensReqCompra.NumIntDoc FROM ItensReqCompra, RequisicaoCompra WHERE RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra AND RequisicaoCompra.Codigo = ? AND ItensReqCompra.Produto = ? AND ItensReqCompra.Fornecedor = ? AND ItensReqCompra.Filial = ? AND ItensReqCompra.Exclusivo = ? AND RequisicaoCompra.FilialEmpresa = ? AND ItensReqCompra.Almoxarifado = ?", lNumIntDoc, objItemReqCompra.lReqCompra, objItemReqCompra.sProduto, objItemReqCompra.lFornecedor, objItemReqCompra.iFilial, objItemReqCompra.iExclusivo, iFilialReqCompra, objItemReqCompra.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 70396
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70397
    
    'Se não encontrou o ItemRC erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70398
    
    objItemReqCompra.lNumIntDoc = lNumIntDoc
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    ItemRC_NumInt_Le = SUCESSO
    
    Exit Function

Erro_ItemRC_NumInt_Le:

    ItemRC_NumInt_Le = gErr
    
    Select Case gErr
    
        Case 70395
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 70396, 70397
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
            
        Case 70398 'ItemRC não cadastrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147740)
    
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Janaina - 18/07 - Estava em GeracaoPedCotacaoOcx
'?? *********************** Corrigir em GeracaoPedCompraAvulsa e outras gerações *********************
'?? Fernando subir
Function TiposProduto_Le_Comprados(ByVal colCod_DescReduzida As AdmColCodigoNome) As Long
'Le todos os Codigos e DescReduzida dos tipos de produto cadastrados

Dim lComando As Long
Dim lErro As Long
Dim tTipoDeProduto As typeTipoDeProduto

On Error GoTo Erro_TiposProduto_Le_Comprados

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 75393

    tTipoDeProduto.sDescricao = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)

    'Seleciona Tipos de Produto
    lErro = Comando_Executar(lComando, "SELECT TipoDeProduto, Descricao FROM TiposDeProduto WHERE Compras=? ORDER BY TipoDeProduto", tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao, PRODUTO_COMPRAVEL)
    If lErro <> AD_SQL_SUCESSO Then gError 75394

    'Lê o primeiro Codigo e DescReduzida
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75395

    'Enquanto Encontrar o Tipo de Produto
    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o Codigo e a DescReduzida lidos na coleção
        colCod_DescReduzida.Add tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao

        'Lê o próximo Codigo e DescReduzida
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75396

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    TiposProduto_Le_Comprados = SUCESSO

    Exit Function

Erro_TiposProduto_Le_Comprados:

    TiposProduto_Le_Comprados = gErr

    Select Case gErr

        Case 75393
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75394
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO1", gErr)

        Case 75395, 75396
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", gErr, tTipoDeProduto.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147741)

    End Select

    'Fecha comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'*****************************************************************************************************
'Janaina - 19/07 - Estava em PedidoCotacaoOcx
'???subir para ClassCOMSelect
Function ItemPedCotacao_Le_CotacaoProduto(ByVal objItemPedCotacao As ClassItemPedCotacao, ByVal objCotacaoProduto As ClassCotacaoProduto) As Long
'LÊ a CotacaoProduto associada ao NumIntDoc do Item do Pedido de Cotacao passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tCotacaoProduto As typeCotacaoProduto
Dim lCotacaoProduto As Long

On Error GoTo Erro_ItemPedCotacao_Le_CotacaoProduto

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76247

    'Busca o ItemPedCotacao  de acordo com NumIntDoc passado
    lErro = Comando_Executar(lComando, "SELECT CotacaoProduto FROM ItemPedCotacao WHERE NumIntDoc=?", lCotacaoProduto, objItemPedCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 76253
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76254
    
    objItemPedCotacao.lCotacaoProduto = lCotacaoProduto
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Fornecedor,Filial FROM CotacaoProduto WHERE NumIntDoc=?", tCotacaoProduto.lNumIntDoc, tCotacaoProduto.lFornecedor, tCotacaoProduto.iFilial, objItemPedCotacao.lCotacaoProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 76248
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76249
    
    'Se nao encontrou==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 76250
    
    objCotacaoProduto.lNumIntDoc = tCotacaoProduto.lNumIntDoc
    objCotacaoProduto.lFornecedor = tCotacaoProduto.lFornecedor
    objCotacaoProduto.iFilial = tCotacaoProduto.iFilial
    
    Call Comando_Fechar(lComando)
    
    ItemPedCotacao_Le_CotacaoProduto = SUCESSO
    
    Exit Function
    
Erro_ItemPedCotacao_Le_CotacaoProduto:

    ItemPedCotacao_Le_CotacaoProduto = SUCESSO
    
    Select Case gErr
    
        Case 76247
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 76248, 76249
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)
            
        Case 76250
            'Erro tratado na rotina chamadora
            
        Case 76253, 76254
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147742)
            
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
'Janaina - 19/07 - Estava em PedidoCotacaoOcx
'???subir para ClassCOMSelect
Function PedidoCotacao_Le_NumIntDoc(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê o Pedido de Cotação apartir do NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tPedidoCotacao As typePedidoCotacao

On Error GoTo Erro_PedidoCotacao_Le_NumIntDoc

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 62864

    With tPedidoCotacao

        .sContato = String(STRING_CONTATO, 0)

        'Seleciona os campos de PedidoCotacao que aparecem na tela onde
        'o Codigo e a FilialEmpresa forem iguais ao do objPedidoCotacao
        lErro = Comando_Executar(lComando, "SELECT  Fornecedor, Filial, Contato, DataEmissao, Data, DataValidade, TipoFrete,CondPagtoPrazo,Codigo,FilialEmpresa FROM PedidoCotacao WHERE NumIntDoc = ? ", .lFornecedor, .iFilial, .sContato, .dtDataEmissao, .dtData, .dtDataValidade, .iTipoFrete, .iCondPagtoPrazo, .lCodigo, .iFilialEmpresa, objPedidoCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 62865

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62866
        If lErro = AD_SQL_SEM_DADOS Then gError 62867

        'Armazena as informações lidas do bd no objeto
        objPedidoCotacao.lCodigo = .lCodigo
        objPedidoCotacao.iFilialEmpresa = .iFilialEmpresa
        objPedidoCotacao.lFornecedor = .lFornecedor
        objPedidoCotacao.iFilial = .iFilial
        objPedidoCotacao.sContato = .sContato
        objPedidoCotacao.dtDataEmissao = .dtDataEmissao
        objPedidoCotacao.dtData = .dtData
        objPedidoCotacao.dtDataValidade = .dtDataValidade
        objPedidoCotacao.iTipoFrete = .iTipoFrete
        objPedidoCotacao.iCondPagtoPrazo = .iCondPagtoPrazo

    End With

    Call Comando_Fechar(lComando)

    PedidoCotacao_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_PedidoCotacao_Le_NumIntDoc:

    PedidoCotacao_Le_NumIntDoc = gErr

    Select Case gErr

        Case 62864
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 62865, 62866
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO1", gErr, objPedidoCotacao.lNumIntDoc)

        Case 62867 'Não existe na tabela de PedidoCotacao

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147743)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Janaina - 19/07 - Estava em GeracaoPedCompraGerCotOcx
Function Requisicoes_Le_Cotacao(ByVal objCotacao As ClassCotacao, ByVal colRequisicoes As Collection) As Long
'Lê as Requisições associadas a Cotação

Dim lErro As Long
Dim lComando As Long
Dim objCotacaoProduto As ClassCotacaoProduto
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim objObservacao As New ClassObservacao
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim iIndice As Integer

On Error GoTo Erro_Requisicoes_Le_Cotacao

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67152
                
    tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
    
    'Lê as Requisições associadas a cotação
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RequisicaoCompra.NumIntDoc, Requisicaocompra.FilialEmpresa, Requisicaocompra.Codigo, RequisicaoCompra.Data, Requisicaocompra.DataLimite, RequisicaoCompra.Urgente, RequisicaoCompra.Requisitante, RequisicaoCompra.Ccl, RequisicaoCompra.TipoDestino, RequisicaoCompra.FornCliDestino, RequisicaoCompra.FilialDestino, RequisicaoCompra.Observacao FROM Requisicaocompra, CotacaoProduto, ItensReqCompra, CotacaoProdutoItemRC, Cotacao WHERE Cotacao.NumIntDoc = CotacaoProduto.Cotacao AND CotacaoProduto.NumIntDoc = CotacaoProdutoItemRC.CotacaoProduto AND CotacaoProdutoItemRC.ItemReqCompra = ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND Cotacao.NumIntDoc = ? AND (ItensReqCompra.Quantidade - ItensReqCompra.QuantCancelada - ItensReqCompra.QuantPedida - ItensReqCompra.QuantRecebida) > 0 ORDER BY RequisicaoCompra.Codigo, RequisicaoCompra.FilialEmpresa", _
    tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sCcl, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67153

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67154
    
    'Se não encontrou nenhuma Requisição, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70441
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
                    
        Set objRequisicaoCompras = New ClassRequisicaoCompras
        
        objRequisicaoCompras.lCodigo = tRequisicaoCompras.lCodigo
        objRequisicaoCompras.iFilialEmpresa = tRequisicaoCompras.iFilialEmpresa
        objRequisicaoCompras.lNumIntDoc = tRequisicaoCompras.lNumIntDoc
        objRequisicaoCompras.dtData = tRequisicaoCompras.dtData
        objRequisicaoCompras.dtDataLimite = tRequisicaoCompras.dtDataLimite
        objRequisicaoCompras.lUrgente = tRequisicaoCompras.lUrgente
        objRequisicaoCompras.lRequisitante = tRequisicaoCompras.lRequisitante
        objRequisicaoCompras.sCcl = tRequisicaoCompras.sCcl
        objRequisicaoCompras.iTipoDestino = tRequisicaoCompras.iTipoDestino
        objRequisicaoCompras.lFornCliDestino = tRequisicaoCompras.lFornCliDestino
        objRequisicaoCompras.iFilialDestino = tRequisicaoCompras.iFilialDestino
    
        'Se a Requisição tem observação
        If tRequisicaoCompras.lObservacao <> 0 Then
              
            objObservacao.lNumInt = tRequisicaoCompras.lObservacao

            'Lê a observação
            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 67155
            If lErro = 53827 Then gError 67156

            objRequisicaoCompras.sObservacao = objObservacao.sObservacao
            
        End If
        
        'Lê os itens de Requisição de Compras
        lErro = CF("ItensReqCompras_Le", objRequisicaoCompras)
        If lErro <> SUCESSO Then gError 67157
        
        For iIndice = objRequisicaoCompras.colItens.Count To 1 Step -1
            
            Set objItemRC = objRequisicaoCompras.colItens(iIndice)
            objItemRC.dQuantComprar = objItemRC.dQuantidade - objItemRC.dQuantCancelada - objItemRC.dQuantPedida - objItemRC.dQuantRecebida
            If objItemRC.dQuantComprar <= 0 Then objRequisicaoCompras.colItens.Remove iIndice
        
        Next
        
        
        'Adiciona nova Requisição de Compras
        colRequisicoes.Add objRequisicaoCompras
        
        'Busca próxima Requisição
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67158
    
    Loop

    Call Comando_Fechar(lComando)

    Requisicoes_Le_Cotacao = SUCESSO

    Exit Function

Erro_Requisicoes_Le_Cotacao:

    Requisicoes_Le_Cotacao = gErr

    Select Case gErr

        Case 67152
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67153, 67154, 67158
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA2", gErr)

        Case 67155, 67157, 70441
        
        Case 67156
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, tRequisicaoCompras.lObservacao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147744)

    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NotasPC_Le(ByVal objNotasPC As ClassNotaPC) As Long
'Lê NotasPC a partir do código.
'Devolve os dados em objNotasPC.

Dim lErro As Long
Dim lComando As Long
Dim sNota As String

On Error GoTo Erro_NotasPC_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103290
    
    sNota = String(STRING_NOTASPC_NOTA, 0)
    
    'Lê os dados da tabela NotasPC
    lErro = Comando_Executar(lComando, "SELECT Nota FROM NotasPC WHERE Codigo = ?", sNota, objNotasPC.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103291

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103292

    'A Nota não existe
    If lErro = AD_SQL_SEM_DADOS Then gError 103293

    'carrega os dados em objNotasPC
    objNotasPC.sNota = sNota

    'Fecha o comando
    Call Comando_Fechar(lComando)

    NotasPC_Le = SUCESSO

    Exit Function

Erro_NotasPC_Le:

    NotasPC_Le = gErr

    Select Case gErr

        Case 103290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103291, 103292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPC", gErr)

        Case 103293
            'Erro tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147745)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NotasPedCompras_Le(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Guarda em objPedidoCompra.colNotasPedCompras todas as Notas de um Pedido

Dim lErro As Long
Dim lComando As Long
Dim sNotaPC As String

On Error GoTo Erro_NotasPedCompras_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103342
    
    sNotaPC = String(STRING_NOTASPC_NOTA, 0)
    
    'Seleciona todas as alcadas da tabela Alcada
    lErro = Comando_Executar(lComando, "SELECT NotaPC FROM NotasPedCompras WHERE NumIntPedido = ? ORDER BY Sequencial", _
    sNotaPC, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 103343

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103344

    Do While lErro = AD_SQL_SUCESSO

        objPedidoCompra.colNotasPedCompras.Add sNotaPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103345

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    NotasPedCompras_Le = SUCESSO

    Exit Function

Erro_NotasPedCompras_Le:

    NotasPedCompras_Le = gErr

    Select Case gErr

        Case 103342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103343, 103344, 103345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPEDCOMPRA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147746)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Moedas_Le_Todas(ByVal colMoedas As Collection) As Long
'Le todas as moedas cadastradas

Dim lErro As Long
Dim lComando As Long
Dim objMoedas As ClassMoedas
Dim sNome As String
Dim sSimbolo As String
Dim iCodigo As Integer
Dim iCodBacen As Integer

On Error GoTo Erro_Moedas_Le_Todas

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103367
    
    sNome = String(STRING_NOME_MOEDA, 0)
    sSimbolo = String(STRING_SIMBOLO_MOEDA, 0)
    
    'Prepara a busca
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, Simbolo, CodBacen FROM Moedas ORDER BY Codigo", iCodigo, sNome, sSimbolo, iCodBacen)
    If lErro <> AD_SQL_SUCESSO Then gError 103368
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103369
    
    'Enquanto encontrar
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objMoedas = New ClassMoedas
        
        objMoedas.iCodigo = iCodigo
        objMoedas.sNome = sNome
        objMoedas.sSimbolo = sSimbolo
        objMoedas.iCodBacen = iCodBacen
        
        'Adiciona a col
        colMoedas.Add objMoedas
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103370
    
    Loop

    Call Comando_Fechar(lComando)

    Moedas_Le_Todas = SUCESSO
    
    Exit Function
    
Erro_Moedas_Le_Todas:

    Moedas_Le_Todas = gErr
    
    Select Case gErr
    
        Case 103367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103368, 103369, 103370
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147747)
    
    End Select

End Function

Public Function Moedas_Le(ByVal objMoedas As ClassMoedas) As Long

Dim lComando As Long
Dim lErro As Long
Dim sNome As String
Dim sSimbolo As String
Dim iPercentual As Integer ' Inserido por Wagner

On Error GoTo Erro_Moedas_Le

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108818

    'Inicializa as strings
    sNome = String(STRING_NOME_MOEDA, 0)
    sSimbolo = String(STRING_SIMBOLO_MOEDA, 0)
    
    '#################################
    'Alterado por Wagner
    'Inserido leitura de ValoresEmPercentual
    'Verifica se existe moeda com o codigo passado
    lErro = Comando_Executar(lComando, "SELECT Nome, Simbolo, ValoresEmPercentual FROM Moedas WHERE Codigo = ?", sNome, sSimbolo, iPercentual, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108819
    '#################################

    'Busca o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108820

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108821

    'Transfere os dados
    objMoedas.sNome = sNome
    objMoedas.sSimbolo = sSimbolo
    
    '#######################
    'Inserido por Wagner
    objMoedas.iPercentual = iPercentual
    '#######################
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Moedas_Le = SUCESSO
    
    Exit Function

Exit Function

Erro_Moedas_Le:

    Moedas_Le = gErr

    Select Case gErr

        Case 108818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108819, 108820
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)

        Case 108821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147748)

    End Select

    Call Comando_Fechar(lComando)

End Function

Function MapaCotacao_ObterPedidos(ByVal objGeracaoMapaCotacao As ClassGeracaoMapaCotacao) As Long
'Preenche a colecao em objMapaCotacao com os pedidos que podem ser faturados e atendem
'aos criterios passados em objGeracaoMapaCotacao com restricoes de No. de pedido, data e categoria ...

Dim lErro As Long
Dim sSelect As String
Dim lComando As Long
Dim lCodigo As Long
Dim sProduto As String
Dim sDescricao As String
Dim dQuantidade As Double
Dim lNumIntItemCotacao As Long
Dim sUM As String
Dim sFornecedor As String
Dim iFilialFornecedor As Integer
Dim sNomeFilialForn As String

On Error GoTo Erro_MapaCotacao_ObterPedidos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 114546

    'Montar SELECT dos pedidos de compra dinamicamente
    Call MapaCotacao_ObterPedidos1(objGeracaoMapaCotacao, sSelect)

    'Inicializar as variaveis de leitura
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    sUM = String(STRING_UM_NOME, 0)
    sNomeFilialForn = String(STRING_FILIAL_FORNECEDOR_NOME, 0)
    
    'Preparar parte fixa do SELECT
    lErro = MapaCotacao_ObterPedidos2(lComando, sSelect, lCodigo, sProduto, sDescricao, dQuantidade, lNumIntItemCotacao, sUM, sFornecedor, iFilialFornecedor, sNomeFilialForn)
    If lErro <> SUCESSO Then gError 114547

    'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT
    lErro = MapaCotacao_ObterPedidos3(lComando, objGeracaoMapaCotacao)
    If lErro <> SUCESSO Then gError 114548

    'Processar os registros lidos
    lErro = MapaCotacao_ObterPedidos4(lComando, objGeracaoMapaCotacao, lCodigo, sProduto, sDescricao, dQuantidade, lNumIntItemCotacao, sUM, sFornecedor, iFilialFornecedor, sNomeFilialForn)
    If lErro <> SUCESSO And lErro <> 114563 Then gError 114549
    
    If lErro = 114563 Then gError 114550
    
    'Fechar Comando
    lErro = Comando_Fechar(lComando)

    MapaCotacao_ObterPedidos = SUCESSO

    Exit Function

Erro_MapaCotacao_ObterPedidos:

    MapaCotacao_ObterPedidos = gErr

    Select Case gErr

        Case 114547 To 114549

        Case 114546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 114550
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO_FILTRO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147749)

    End Select

    Call Comando_Fechar(lComando)

End Function

Sub MapaCotacao_ObterPedidos1(objGeracaoMapaCotacao As ClassGeracaoMapaCotacao, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String, sGroupBy As String
Dim iIndice As Integer

    sFields = "PedidoCotacaoN.Codigo AS CodItem, Produtos.Codigo AS Produto, Produtos.Descricao, CotacaoProdutoN.Quantidade, ItemPedCotacaoN.NumIntDoc AS NumIntItemPedCotacao, CotacaoProdutoN.UM, Fornecedores.NomeReduzido AS Fornecedor, FiliaisFornecedores.CodFilial AS CodFilialForn, FiliaisFornecedores.Nome AS NomeFilialForn"
    sFrom = " FROM PedidoCotacaoN, ItemPedCotacaoN, Produtos, CotacaoProdutoN, Fornecedores, FiliaisFornecedores "
    
    
    sWhere = " WHERE  PedidoCotacaoN.FilialEmpresa = ? AND PedidoCotacaoN.numIntDoc = ItemPedCotacaoN.PedidoCotacao AND ItemPedCotacaoN.CotacaoProduto = CotacaoProdutoN.NumIntDoc AND CotacaoProdutoN.Produto = Produtos.Codigo  AND PedidoCotacaoN.Fornecedor = Fornecedores.Codigo AND Fornecedores.Codigo = FiliaisFornecedores.CodFornecedor  AND PedidoCotacaoN.Filial = FiliaisFornecedores.CodFilial AND ( PedidoCotacaoN.DataValidade >= ? OR PedidoCotacaoN.DataValidade = ?) AND PedidoCotacaoN.Status <> ? "
    sGroupBy = " GROUP BY Produtos.Codigo, PedidoCotacaoN.Codigo, Produtos.Descricao, CotacaoProdutoN.Quantidade, ItemPedCotacaoN.NumIntDoc, Fornecedores.NomeReduzido, CotacaoProdutoN.UM, FiliaisFornecedores.CodFilial, FiliaisFornecedores.Nome "
    sOrderBy = " ORDER BY PedidoCotacaoN.Codigo, Produtos.Codigo "
    
    If objGeracaoMapaCotacao.ColItensCategoria.Count > 0 Then
        sFrom = sFrom & ", ProdutoCategoria "
        sWhere = sWhere & " AND Produtos.Codigo = ProdutoCategoria.Produto "
    End If
    
    'Verifica a parte dinamica e adicona ao select se necessário ...
    If objGeracaoMapaCotacao.lCodigoDe > 0 Then sWhere = sWhere & " AND PedidoCotacaoN.Codigo >= ?"
    If objGeracaoMapaCotacao.lCodigoAte > 0 Then sWhere = sWhere & " AND PedidoCotacaoN.Codigo <= ?"
    
    If objGeracaoMapaCotacao.dtDataDe > 0 Then sWhere = sWhere & " AND PedidoCotacaoN.Data >= ?"
    If objGeracaoMapaCotacao.dtDataAte > 0 Then sWhere = sWhere & " AND PedidoCotacaoN.Data <= ?"
    
    If Len(Trim(objGeracaoMapaCotacao.sCategoria)) > 0 Then
    
        sWhere = sWhere & " AND (ProdutoCategoria.Categoria = '" & objGeracaoMapaCotacao.sCategoria & "' "
        
        For iIndice = 1 To objGeracaoMapaCotacao.ColItensCategoria.Count
            
            If iIndice = 1 Then
                sWhere = sWhere & " AND (ProdutoCategoria.Item = '" & objGeracaoMapaCotacao.ColItensCategoria(iIndice) & "'"
            Else
                sWhere = sWhere & " OR ProdutoCategoria.Item = '" & objGeracaoMapaCotacao.ColItensCategoria(iIndice) & "'"
            End If
            
        Next
        
        If iIndice >= 1 Then sWhere = sWhere & "))"
        
    End If
    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sGroupBy & sOrderBy

End Sub

 Function MapaCotacao_ObterPedidos2(ByVal lComando As Long, ByVal sSelect As String, vlCodigo As Variant, vsProduto As Variant, vsDescricao As Variant, vdQuantidade As Variant, vlNumIntItemCotacao As Variant, vsUM As Variant, vsFornecedor As Variant, viFilialFornecedor As Variant, vsNomeFilialForn As Variant) As Long
'Preparar parte fixa do SELECT

Dim ret As Integer, lErro As Long

On Error GoTo Erro_MapaCotacao_ObterPedidos2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 114551

    ret = Comando_BindVarInt(lComando, vlCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 114552

    ret = Comando_BindVarInt(lComando, vsProduto)
    If (ret <> AD_SQL_SUCESSO) Then gError 114553

    ret = Comando_BindVarInt(lComando, vsDescricao)
    If (ret <> AD_SQL_SUCESSO) Then gError 114554

    ret = Comando_BindVarInt(lComando, vdQuantidade)
    If (ret <> AD_SQL_SUCESSO) Then gError 114555
    
    ret = Comando_BindVarInt(lComando, vlNumIntItemCotacao)
    If (ret <> AD_SQL_SUCESSO) Then gError 114585
    
    ret = Comando_BindVarInt(lComando, vsUM)
    If (ret <> AD_SQL_SUCESSO) Then gError 114631
    
    ret = Comando_BindVarInt(lComando, vsFornecedor)
    If (ret <> AD_SQL_SUCESSO) Then gError 114632
    
    ret = Comando_BindVarInt(lComando, viFilialFornecedor)
    If (ret <> AD_SQL_SUCESSO) Then gError 114633
    
    ret = Comando_BindVarInt(lComando, vsNomeFilialForn)
    If (ret <> AD_SQL_SUCESSO) Then gError 114634
    
    MapaCotacao_ObterPedidos2 = SUCESSO

    Exit Function

Erro_MapaCotacao_ObterPedidos2:

    MapaCotacao_ObterPedidos2 = gErr

    Select Case gErr

        Case 114551 To 114555, 114585, 114631 To 114634
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147750)

    End Select

End Function

Function MapaCotacao_ObterPedidos3(ByVal lComando As Long, ByVal objGeracaoMapaCotacao As ClassGeracaoMapaCotacao) As Long
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT

Dim lErro As Long
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant
Dim vsCategoria As Variant
Dim vsItemCategoria As Variant
Dim iIndice As Integer
Dim vdtDataHoje As Variant
Dim vdtDataNula As Variant
Dim viStatus As Variant, viFilialEmpresa As Variant

On Error GoTo Erro_MapaCotacao_ObterPedidos3

    viFilialEmpresa = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 114650
    
    vdtDataHoje = gdtDataHoje
    lErro = Comando_BindVarInt(lComando, vdtDataHoje)
    If (lErro <> AD_SQL_SUCESSO) Then gError 114650
    
    vdtDataNula = DATA_NULA
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 114650
    
    viStatus = STATUS_ATUALIZADO
    lErro = Comando_BindVarInt(lComando, viStatus)
    If (lErro <> AD_SQL_SUCESSO) Then gError 114651
    
    If objGeracaoMapaCotacao.lCodigoDe <> 0 Then
        vlCodigoDe = objGeracaoMapaCotacao.lCodigoDe
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 114555
    End If

    If objGeracaoMapaCotacao.lCodigoAte <> 0 Then
        vlCodigoAte = objGeracaoMapaCotacao.lCodigoAte
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 114556
    End If

    If objGeracaoMapaCotacao.dtDataDe <> DATA_NULA Then
        vdtDataDe = objGeracaoMapaCotacao.dtDataDe
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 114557
    End If

    If objGeracaoMapaCotacao.dtDataAte <> DATA_NULA Then
        vdtDataAte = objGeracaoMapaCotacao.dtDataAte
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 114558
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 114561

    MapaCotacao_ObterPedidos3 = SUCESSO

    Exit Function

Erro_MapaCotacao_ObterPedidos3:

    MapaCotacao_ObterPedidos3 = gErr

    Select Case gErr

        Case 114555 To 114561, 114650, 114651
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147751)

    End Select

End Function

Function MapaCotacao_ObterPedidos4(ByVal lComando As Long, ByVal objGeracaoMapaCotacao As ClassGeracaoMapaCotacao, lCodigo As Long, sProduto As String, sDescricao As String, dQuantidade As Double, lNumIntItemCotacao As Long, sUM As String, sFornecedor As String, iFilialFornecedor As Integer, sNomeFilialForn As String) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long
Dim objItemMapaCotacao As ClassMapaCotacaoItemCotacao

On Error GoTo Erro_MapaCotacao_ObterPedidos4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114562

    'Se não encontrou==>Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 114563

    Do While lErro = AD_SQL_SUCESSO

        Set objItemMapaCotacao = New ClassMapaCotacaoItemCotacao
        
        objItemMapaCotacao.lPedidoCotacao = lCodigo
        objItemMapaCotacao.sProduto = sProduto
        objItemMapaCotacao.sDescricao = sDescricao
        objItemMapaCotacao.dQuantidade = dQuantidade
        objItemMapaCotacao.lNumIntItemCotacao = lNumIntItemCotacao
        objItemMapaCotacao.sUM = sUM
        objItemMapaCotacao.sFornecedor = sFornecedor
        objItemMapaCotacao.iFilialForn = iFilialFornecedor
        objItemMapaCotacao.sNomeFilialForn = sNomeFilialForn
        
        'Adiciona objItemMapaCotacao na coleção de Pedidos de Cotacao
        objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao.Add objItemMapaCotacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114566

    Loop

    MapaCotacao_ObterPedidos4 = SUCESSO

    Exit Function

Erro_MapaCotacao_ObterPedidos4:

    MapaCotacao_ObterPedidos4 = gErr

    Select Case gErr

        Case 114562, 114566
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", gErr)

        Case 114563
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147752)

    End Select

End Function

Function MapaCotacao_Le(ByVal objGeracaoMapaCotacao As ClassGeracaoMapaCotacao) As Long
'Le o mapa de cotacao e os itens selecionados

Dim lErro As Long
Dim sSelect As String
Dim lComando As Long
Dim tItemMapaCotacao As typeItemMapaCotacao
Dim objItemMapaCotacao As ClassMapaCotacaoItemCotacao

On Error GoTo Erro_MapaCotacao_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 114580

    'Busca a ser executada ...
    sSelect = "SELECT CodItem, Produto, Descricao, Quantidade, NumIntItemPedCotacao, Codigo, Data, TaxaFinanceira, FilialEmpresa, Fornecedor, UM, CodFilialForn, NomeFilialForn FROM MapaCotacaoTodos WHERE Codigo = ? ORDER BY CodItem, Produto"

    'Prepara a execucao do select
    With tItemMapaCotacao

        .sProduto = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
        .sUM = String(STRING_UM_NOME, 0)
        .sNomeFilialForn = String(STRING_FILIAL_FORNECEDOR_NOME, 0)

        lErro = Comando_Executar(lComando, sSelect, .lPedCotacao, .sProduto, .sDescricao, .dQuantidade, .lNumIntItemPedCotacao, .lMapaCotacao, .dtData, .dTaxaFinanceira, .iFilialEmpresa, .sFornecedor, .sUM, .iFilialForn, .sNomeFilialForn, objGeracaoMapaCotacao.objMapaCotacao.lCodigo)

    End With
    If lErro <> AD_SQL_SUCESSO Then gError 114581

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114582

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 114583
    
    objGeracaoMapaCotacao.objMapaCotacao.dtData = tItemMapaCotacao.dtData
    objGeracaoMapaCotacao.objMapaCotacao.dTaxaFinanceira = tItemMapaCotacao.dTaxaFinanceira

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemMapaCotacao = New ClassMapaCotacaoItemCotacao

        With tItemMapaCotacao

            objItemMapaCotacao.dQuantidade = .dQuantidade
            objItemMapaCotacao.sProduto = .sProduto
            objItemMapaCotacao.sDescricao = .sDescricao
            objItemMapaCotacao.lPedidoCotacao = .lPedCotacao
            objItemMapaCotacao.lNumIntItemCotacao = .lNumIntItemPedCotacao
            objItemMapaCotacao.iFilialEmpresa = .iFilialEmpresa
            objItemMapaCotacao.sFornecedor = .sFornecedor
            objItemMapaCotacao.sUM = .sUM
            objItemMapaCotacao.iFilialForn = .iFilialForn
            objItemMapaCotacao.sNomeFilialForn = .sNomeFilialForn

        End With
        
        objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao.Add objItemMapaCotacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114584

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    MapaCotacao_Le = SUCESSO

    Exit Function

Erro_MapaCotacao_Le:

    MapaCotacao_Le = gErr

    Select Case gErr

        Case 114580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 114581, 114582, 114584
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOTACAO", gErr)

        Case 114583

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147753)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 03/03/04
Public Function RelABCCompras_Gera(ByVal objRelABCComprasTela As ClassRelABCComprasTela, ByVal colItensRelABCCompras As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_RelABCCompras_Gera

    'Lê os dados que farão parte do relatório
    lErro = SldDiaEst_Le_RelABCCompras(objRelABCComprasTela, colItensRelABCCompras)
    If lErro <> SUCESSO And lErro <> 127027 Then gError 127069
    
    'se não encontrou dados
    If lErro = 127027 Then gError 127081

    'Insere os registro na tabela auxiliar RelABCCompras
    lErro = CF("RelABCCompras_Grava", colItensRelABCCompras)
    If lErro <> SUCESSO Then gError 125908

    RelABCCompras_Gera = SUCESSO
    
    Exit Function

Erro_RelABCCompras_Gera:

    RelABCCompras_Gera = gErr
    
    Select Case gErr
    
        Case 125908, 127069
        
        Case 127081
            Call Rotina_Erro(vbOKOnly, "ERRO_DADOS_REL_NAO_ENCONTRADO", gErr, Error$)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147754)
    
    End Select

End Function

'Incluída por Luiz Nogueira em 03/03/04
Public Function SldDiaEst_Le_RelABCCompras(ByVal objRelABCComprasTela As ClassRelABCComprasTela, ByVal colItensRelABCCompras As Collection) As Long
'Função que monta o Select que será executado posteriormente

Dim lErro As Long
Dim sSelect As String
Dim tRelABCCompras As TypeRelABCComprasVar
Dim lComando As Long

On Error GoTo Erro_SldDiaEst_Le_RelABCCompras

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 127004

    'Monta o select que será executado
    lErro = SldDiaEst_Le_RelABCCompras1(objRelABCComprasTela, sSelect)
    If lErro <> SUCESSO Then gError 127005

    'Prepara as variáveis que receberão o resultado do select
    lErro = SldDiaEst_Le_RelABCCompras2(lComando, tRelABCCompras, objRelABCComprasTela, sSelect)
    If lErro <> SUCESSO Then gError 127006

    'Prepara os filtros que serão passados para o select e executa o select no BD
    lErro = SldDiaEst_Le_RelABCCompras3(lComando, tRelABCCompras, objRelABCComprasTela, sSelect)
    If lErro <> SUCESSO Then gError 127007

    'Recebe os dados lidos e armazena-os na coleção
    lErro = SldDiaEst_Le_RelABCCompras4(lComando, tRelABCCompras, colItensRelABCCompras, sSelect)
    If lErro <> SUCESSO And lErro <> 127026 Then gError 127008
    
    'Se não encontrou dados => erro
    If lErro = 127026 Then gError 127027
    
    'Calcula o percentual de participação de cada produto
    lErro = RelsABC_Calcula_PercParticipacao(colItensRelABCCompras)
    If lErro <> SUCESSO Then gError 127028

    'fecha o comando
    Call Comando_Fechar(lComando)

    SldDiaEst_Le_RelABCCompras = SUCESSO

    Exit Function

Erro_SldDiaEst_Le_RelABCCompras:

    SldDiaEst_Le_RelABCCompras = gErr

    Select Case gErr

        Case 127004
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 127005 To 127008, 127028

        Case 127027 'sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147755)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaEst_Le_RelABCCompras1(ByVal objRelABCComprasTela As ClassRelABCComprasTela, sSelect As String) As Long
'Monta o select que será executado

Dim lErro As Long
Dim sFrom As String
Dim sWhere As String
Dim sGroupBy As String
Dim sHaving As String
Dim sOrderBy As String
Dim iWhere As Integer
Dim iIndice As Integer

On Error GoTo Erro_SldDiaEst_Le_RelABCCompras1

    'select fixo
    sSelect = "SELECT SldDiaEst.Produto, SUM(SldDiaEst.QuantComp), SUM(SldDiaEst.ValorComp) "

    'From Fixo
    sFrom = "FROM SldDiaEst "
    
    'Cláusula Group By
    sGroupBy = "GROUP BY SldDiaEst.Produto "
    
    'Cláusula Having
    sHaving = "HAVING (SUM(SldDiaEst.QuantComp)>0 OR SUM(SldDiaEst.ValorComp)>0) "
    
    'Cláusula Order By
    sOrderBy = "ORDER BY SUM(SldDiaEst.ValorComp) DESC"

    '*** DATA ***
    'Verifica se a Data De é diferente de data Nula se for  será utilizada como Filtro
    If objRelABCComprasTela.dtDataDe <> DATA_NULA Then

        'se já existe cláusula where
        If iWhere = 1 Then
            sWhere = sWhere & " AND SldDiaEst.Data >= ? "
        Else
            sWhere = "WHERE SldDiaEst.Data >= ? "
            iWhere = 1
        End If

    End If

    'Verifica se a Data Até é diferente de data Nula se for  será utilizada como Filtro
    If objRelABCComprasTela.dtDataAte <> DATA_NULA Then

        'se já existe cláusula where
        If iWhere = 1 Then
            sWhere = sWhere & " AND SldDiaEst.Data <= ? "
        Else
            sWhere = "WHERE SldDiaEst.Data <= ? "
            iWhere = 1
        End If
        
    End If
    '*** DATA ***

    '*** FILIALEMPRESA ***
    'Verifica se será utilizado o filtro FilialEmpresaDe
    If objRelABCComprasTela.iFilialEmpresaDe > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro FilialEmpresa>= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaEst.FilialEmpresa >=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro FilialEmpresa>=
            sWhere = "WHERE SldDiaEst.FilialEmpresa >=? "
            iWhere = 1
        
        End If

    End If
            
    'Verifica se será utilizado o filtro FilialEmpresaAte
    If objRelABCComprasTela.iFilialEmpresaAte > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro FilialEmpresa<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaEst.FilialEmpresa<=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro FilialEmpresa<=
            sWhere = "WHERE SldDiaEst.FilialEmpresa<=? "
        
        End If

    End If
    '*** FILIALEMPRESA ***
            
    '*** PRODUTO ***
    'Verifica se será utilizado o filtro de produtode
    If Len(Trim(objRelABCComprasTela.sProdutoDe)) > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaEst.Produto>=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaEst.Produto>=? "
            iWhere = 1
        
        End If
    
    End If
    
    'Verifica se será utilizado o filtro de produtoate
    If Len(Trim(objRelABCComprasTela.sProdutoAte)) > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaEst.Produto<=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaEst.Produto<=? "
            iWhere = 1
        
        End If
    
    End If
    '*** PRODUTO ***
            
    '*** PRODUTOSTOP ***
    'Se é para selecionar apenas os xx Produtos mais comprados
    If objRelABCComprasTela.iProdutosTop > 0 Then
    
        'altera a cláusula select para pegar apenas os TOP xx produtos
        sSelect = Mid(sSelect, 1, 6) & " TOP " & objRelABCComprasTela.iProdutosTop & Mid(sSelect, 7)
    
    End If
    '*** PRODUTOSTOP ***
    
    '*** TIPODEPRODUTO ***
    'Verifica se será utilizado filtro por tipo de produto
    If objRelABCComprasTela.iTipoProduto > 0 Then

        'adiciona a tabela Produtos à cláusula From
        sFrom = sFrom & ", Produtos "
        
        'se já existe cláusula where
        If iWhere = 1 Then
            'adiciona cláusula para linkar as tabelas SldDiaEst e Produtos
            sWhere = sWhere & " AND SldDiaEst.Produto = Produtos.Codigo AND Produtos.Tipo=? "
        
        'se não existe
        Else
            'cria cláusula para linkar as tabelas SldDiaEst e Produtos
            sWhere = "WHERE SldDiaEst.Produto = Produtos.Codigo AND Produtos.Tipo=? "
            iWhere = 1
        End If
        
    End If
    '*** TIPODEPRODUTO ***
    
    '*** CATEGORIA DE PRODUTO ***
    'Verifica se será utilizado filtro por categoria de produto
    If Len(Trim(objRelABCComprasTela.sCategoria)) <> 0 Then

        'Altera o select para ler o item de categoria associado ao produto
        sSelect = sSelect & ", MAX(ProdutoCategoria.Item) "
        
        'adiciona a tabela Produtos à cláusula From
        sFrom = sFrom & ", ProdutoCategoria "
        
        'se já existe cláusula where
        If iWhere = 1 Then
            'adiciona cláusula para linkar as tabelas SldDiaEst e Produtos e filtrar por categoria e item de categoria
            sWhere = sWhere & " AND SldDiaEst.Produto=ProdutoCategoria.Produto AND ProdutoCategoria.Categoria=? AND (ProdutoCategoria.Item=? "
        
        'se não existe
        Else
            'cria cláusula para linkar as tabelas SldDiaEst e Produtos e filtrar por categoria e item de categoria
            sWhere = "WHERE SldDiaEst.Produto=ProdutoCategoria.Produto AND ProdutoCategoria.Categoria=? AND (ProdutoCategoria.Item=? "
            iWhere = 1
        End If
        
        'Para cada item de categoria selecionado (o loop começa com índice 2, pois o filtro para primeira categoria já foi adicionado acima
        For iIndice = 2 To objRelABCComprasTela.ColItensCategoria.Count
            'Inclui o filtro na cláusula where
            sWhere = sWhere & "OR ProdutoCategoria.Item=? "
        Next
        
        'Fecha o parênteses dos itens de categoria na cláusula WHERE
        sWhere = sWhere & ") "

    End If
    '*** CATEGORIA DE PRODUTO ***
    
    'Monta o select que será executado
    sSelect = sSelect & sFrom & sWhere & sGroupBy & sHaving & sOrderBy
    
    SldDiaEst_Le_RelABCCompras1 = SUCESSO

    Exit Function

Erro_SldDiaEst_Le_RelABCCompras1:

    SldDiaEst_Le_RelABCCompras1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147756)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaEst_Le_RelABCCompras2(ByVal lComando As Long, tRelABCCompras As TypeRelABCComprasVar, ByVal objRelABCComprasTelas As ClassRelABCComprasTela, sSelect As String) As Long
'Função que faz o bind das variáveis que serão lidas no select

Dim lErro As Long

On Error GoTo Erro_SldDiaEst_Le_RelABCCompras2

    With tRelABCCompras

        .vsProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_PrepararInt(lComando, sSelect)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127009

        lErro = Comando_BindVarInt(lComando, .vsProduto)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127010

        .vdQuantidade = CDbl(.vdQuantidade)
        lErro = Comando_BindVarInt(lComando, .vdQuantidade)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127011

        .vdValor = CDbl(.vdValor)
        lErro = Comando_BindVarInt(lComando, .vdValor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127012
        
        'Se foi passada uma categoria como filtro
        If Len(Trim(objRelABCComprasTelas.sCategoria)) > 0 Then
            
            'inicializa a variável que receberá o item de categoria
            .vsItemCategoria = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
            
            .vsItemCategoria = CStr(.vsItemCategoria)
            lErro = Comando_BindVarInt(lComando, .vsItemCategoria)
            If (lErro <> AD_SQL_SUCESSO) Then gError 127034
        
        End If
        
    End With

    SldDiaEst_Le_RelABCCompras2 = SUCESSO

    Exit Function

Erro_SldDiaEst_Le_RelABCCompras2:

    SldDiaEst_Le_RelABCCompras2 = gErr

    Select Case gErr

        Case 127009 To 127012, 127034
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147757)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaEst_Le_RelABCCompras3(ByVal lComando As Long, tRelABCCompras As TypeRelABCComprasVar, ByVal objRelABCComprasTela As ClassRelABCComprasTela, sSelect As String)
'Faz o bind dos filtros passados para o select e executa o select no BD

Dim lErro As Long
Dim iIndice As Long
Dim sItemCategoria As String

On Error GoTo Erro_SldDiaEst_Le_RelABCCompras3

    'Se foi passado o filtro DataDe
    If objRelABCComprasTela.dtDataDe <> DATA_NULA Then

        tRelABCCompras.vdtDataDe = CDate(objRelABCComprasTela.dtDataDe)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127013

    End If

    'Se foi passado o filtro DataAte
    If objRelABCComprasTela.dtDataAte <> DATA_NULA Then

        tRelABCCompras.vdtDataAte = CDate(objRelABCComprasTela.dtDataAte)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127014

    End If

    'Se foi passado filtro FilialEmpresaDe
    If objRelABCComprasTela.iFilialEmpresaDe > 0 Then

        tRelABCCompras.viFilialEmpresaDe = CInt(objRelABCComprasTela.iFilialEmpresaDe)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.viFilialEmpresaDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127015


    End If

    'Se foi passado filtro FilialEmpresaAte
    If objRelABCComprasTela.iFilialEmpresaAte > 0 Then

        tRelABCCompras.viFilialEmpresaAte = CInt(objRelABCComprasTela.iFilialEmpresaAte)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.viFilialEmpresaAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127016


    End If
    
    'Verifica se foi passado o filtro ProdutoDe
    If Len(Trim(objRelABCComprasTela.sProdutoDe)) > 0 Then

        tRelABCCompras.vsProdutoDe = CStr(objRelABCComprasTela.sProdutoDe)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.vsProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127017

    End If

    'Verifica se foi passado o filtro ProdutoAte
    If Len(Trim(objRelABCComprasTela.sProdutoAte)) > 0 Then

        tRelABCCompras.vsProdutoAte = CStr(objRelABCComprasTela.sProdutoAte)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.vsProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127018

    End If
    
    'Verifica se foi passado o filtro TipoProdutoDe
    If objRelABCComprasTela.iTipoProduto <> 0 Then

        tRelABCCompras.viTipoProduto = CInt(objRelABCComprasTela.iTipoProduto)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.viTipoProduto)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127019

    End If

    'Se foi passado o filtro Categoria
    If Len(Trim(objRelABCComprasTela.sCategoria)) <> 0 Then

        tRelABCCompras.vsCategoria = CStr(objRelABCComprasTela.sCategoria)
        lErro = Comando_BindVarInt(lComando, tRelABCCompras.vsCategoria)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127021

    End If
    
    'Redimensiona o array que armazena os itens de categoria a serem usados como filtro
    ReDim tRelABCCompras.avsItensCategoria(objRelABCComprasTela.ColItensCategoria.Count)
    
    'Para cada item de categoria
    For iIndice = 1 To objRelABCComprasTela.ColItensCategoria.Count
    
        'Guarda na variável o item de categoria que está na coleção
        sItemCategoria = objRelABCComprasTela.ColItensCategoria(iIndice)
        
        'Se o ItemCategoria está preenchido
        If Len(Trim(sItemCategoria)) <> 0 Then
    
            tRelABCCompras.avsItensCategoria(iIndice) = CStr(sItemCategoria)
            lErro = Comando_BindVarInt(lComando, tRelABCCompras.avsItensCategoria(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then gError 127022
    
        End If
    
    Next
        
    'Executa a leitura no BD
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 127023

    SldDiaEst_Le_RelABCCompras3 = SUCESSO

    Exit Function

Erro_SldDiaEst_Le_RelABCCompras3:

    SldDiaEst_Le_RelABCCompras3 = gErr

    Select Case gErr

        Case 127013 To 127023
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147758)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaEst_Le_RelABCCompras4(ByVal lComando As Long, tRelABCCompras As TypeRelABCComprasVar, ByVal colItensRelABCCompras As Collection, ByVal sSelect As String) As Long
'Busca os dados lidos no BD e armazena-os em colItensRelABCCompras

Dim lErro As Long
Dim objRelABCCompras As ClassRelABCCompras
Dim lRanking As Long

On Error GoTo Erro_SldDiaEst_Le_RelABCCompras4

    'Busca o primeiro registro lido no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 127024

    'Se não encontrou dados => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 127026

    'Enquanto houver dados no BD
    Do While lErro = AD_SQL_SUCESSO

        'Incrementa a variável que controla a posição do produto no ranking
        lRanking = lRanking + 1

        'Instancia um novo obj
        Set objRelABCCompras = New ClassRelABCCompras

        'Guarda os dados lidos no obj
        With objRelABCCompras
            .sProduto = tRelABCCompras.vsProduto
            .lRanking = lRanking
            .dQuantidade = tRelABCCompras.vdQuantidade
            .dValor = tRelABCCompras.vdValor
            .sItemCategoria = tRelABCCompras.vsItemCategoria
        End With

        'Adiciona o obj à coleção
        colItensRelABCCompras.Add objRelABCCompras

        'Busca o próximo registro lido
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 127025

    Loop

    SldDiaEst_Le_RelABCCompras4 = SUCESSO

    Exit Function

Erro_SldDiaEst_Le_RelABCCompras4:

    SldDiaEst_Le_RelABCCompras4 = gErr

    Select Case gErr

        Case 127024 To 127025
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case 127026 'SEM DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147759)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Public Function RelABCFornecedores_Gera(ByVal objRelABCFornecedoresTela As ClassRelABCFornecedoresTela, ByVal colItensRelABCFornecedores As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_RelABCFornecedores_Gera

    'Lê os dados que farão parte do relatório
    lErro = SldDiaForn_Le_RelABCFornecedores(objRelABCFornecedoresTela, colItensRelABCFornecedores)
    If lErro <> SUCESSO And lErro <> 127041 Then gError 127070
    
    'se não encontrou dados p/o relatório
    If lErro = 127041 Then gError 127082

    'Insere os registro na tabela auxiliar RelABCCompras
    lErro = CF("RelABCFornecedores_Grava", colItensRelABCFornecedores)
    If lErro <> SUCESSO Then gError 127071

    RelABCFornecedores_Gera = SUCESSO
    
    Exit Function

Erro_RelABCFornecedores_Gera:

    RelABCFornecedores_Gera = gErr
    
    Select Case gErr
    
        Case 127070, 127071
    
        Case 127082
            Call Rotina_Erro(vbOKOnly, "ERRO_DADOS_REL_NAO_ENCONTRADO", gErr, Error$)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147760)
    
    End Select

End Function

'Incluída por Luiz Nogueira em 03/03/04
Public Function SldDiaForn_Le_RelABCFornecedores(ByVal objRelABCFornecedoresTela As ClassRelABCFornecedoresTela, ByVal colItensRelABCFornecedores As Collection) As Long
'Lê em SldDiaForn os dados necessários para o relatório ABC de Fornecedores

Dim lErro As Long
Dim sSelect As String
Dim tRelABCFornecedores As TypeRelABCFornecedoresVar
Dim lComando As Long

On Error GoTo Erro_SldDiaForn_Le_RelABCFornecedores

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 127035

    'Função que Monta o select
    lErro = SldDiaForn_Le_RelABCFornecedores1(objRelABCFornecedoresTela, sSelect)
    If lErro <> SUCESSO Then gError 127036

    'Função que Prepara a Parte Fixa do Select
    lErro = SldDiaForn_Le_RelABCFornecedores2(lComando, tRelABCFornecedores, objRelABCFornecedoresTela, sSelect)
    If lErro <> SUCESSO Then gError 127037

    'Prepara os parâmetros que variam de acordo com a seleção do usuário
    lErro = SldDiaForn_Le_RelABCFornecedores3(lComando, tRelABCFornecedores, objRelABCFornecedoresTela, sSelect)
    If lErro <> SUCESSO Then gError 127038

    'Função que preenche a Coleção com os Registros retornados pelo Select
    lErro = SldDiaForn_Le_RelABCFornecedores4(lComando, tRelABCFornecedores, colItensRelABCFornecedores, sSelect)
    If lErro <> SUCESSO And lErro <> 127062 Then gError 127039
    
    'Se não encontrou dados => erro
    If lErro = 127062 Then gError 127041
    
    'Calcula o percentual de participação de cada produto
    lErro = RelsABC_Calcula_PercParticipacao(colItensRelABCFornecedores)
    If lErro <> SUCESSO Then gError 127040

    'fecha o comando
    Call Comando_Fechar(lComando)

    SldDiaForn_Le_RelABCFornecedores = SUCESSO

    Exit Function

Erro_SldDiaForn_Le_RelABCFornecedores:

    SldDiaForn_Le_RelABCFornecedores = gErr

    Select Case gErr

        Case 127035
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 127036 To 127040

        Case 127041 'sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147761)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaForn_Le_RelABCFornecedores1(ByVal objRelABCFornecedoresTela As ClassRelABCFornecedoresTela, sSelect As String) As Long
'Monta o select que será executado

Dim lErro As Long
Dim sFrom As String
Dim sWhere As String
Dim sGroupBy As String
Dim sHaving As String
Dim sOrderBy As String
Dim iWhere As Integer
Dim iIndice As Integer

On Error GoTo Erro_SldDiaForn_Le_RelABCFornecedores1

    'select fixo
    sSelect = "SELECT SldDiaForn.Fornecedor, SUM(SldDiaForn.ValorCompra) "

    'From Fixo
    sFrom = "FROM SldDiaForn "
    
    'Cláusula Group By
    sGroupBy = "GROUP BY SldDiaForn.Fornecedor "
    
    'Cláusula Having
    sHaving = "HAVING SUM(SldDiaForn.ValorCompra)>0 "
    
    'Cláusula Order By
    sOrderBy = "ORDER BY SUM(SldDiaForn.ValorCompra) DESC"

    '*** DATA ***
    'Verifica se a Data De é diferente de data Nula se for  será utilizada como Filtro
    If objRelABCFornecedoresTela.dtDataDe <> DATA_NULA Then

        'se já existe cláusula where
        If iWhere = 1 Then
            sWhere = sWhere & " AND SldDiaForn.Data >= ? "
        Else
            sWhere = "WHERE SldDiaForn.Data >= ? "
            iWhere = 1
        End If

    End If

    'Verifica se a Data Até é diferente de data Nula se for  será utilizada como Filtro
    If objRelABCFornecedoresTela.dtDataAte <> DATA_NULA Then

        'se já existe cláusula where
        If iWhere = 1 Then
            sWhere = sWhere & " AND SldDiaForn.Data <= ? "
        Else
            sWhere = "WHERE SldDiaForn.Data <= ? "
            iWhere = 1
        End If
        
    End If
    '*** DATA ***

    '*** FILIALEMPRESA ***
    'Verifica se será utilizado o filtro FilialEmpresaDe
    If objRelABCFornecedoresTela.iFilialEmpresaDe > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro FilialEmpresa>= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.FilialEmpresa >=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro FilialEmpresa>=
            sWhere = "WHERE SldDiaForn.FilialEmpresa >=? "
        
        End If

    End If
            
    'Verifica se será utilizado o filtro FilialEmpresaAte
    If objRelABCFornecedoresTela.iFilialEmpresaAte > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro FilialEmpresa<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.FilialEmpresa<=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro FilialEmpresa<=
            sWhere = "WHERE SldDiaForn.FilialEmpresa<=? "
        
        End If

    End If
    '*** FILIALEMPRESA ***
            
    '*** PRODUTO ***
    'Verifica se será utilizado o filtro de produtode
    If Len(Trim(objRelABCFornecedoresTela.sProdutoDe)) > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.Produto>=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaForn.Produto>=? "
            iWhere = 1
        
        End If
    
    End If
    
    'Verifica se será utilizado o filtro de produtoate
    If Len(Trim(objRelABCFornecedoresTela.sProdutoAte)) > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.Produto<=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaForn.Produto<=? "
            iWhere = 1
        
        End If
    
    End If
    '*** PRODUTO ***
    
    '*** FORNECEDOR ***
    'Verifica se será utilizado o filtro de FornecedorDe
    If objRelABCFornecedoresTela.lFornecedorDe > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.Fornecedor>=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaForn.Fornecedor>=? "
        
        End If
    
    End If
    
    'Verifica se será utilizado o filtro de FornecedorAte
    If objRelABCFornecedoresTela.lFornecedorAte > 0 Then
    
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona o filtro Produto<= à cláusula WHERE
            sWhere = sWhere & "AND SldDiaForn.Fornecedor<=? "
        
        'senão
        Else
        
            'cria a cláusula where com o filtro Produto<=
            sWhere = "WHERE SldDiaForn.Fornecedor<=? "
        
        End If
    
    End If
    '*** FORNECEDOR ***
            
    '*** TIPODEPRODUTO ***
    'Verifica se será utilizado filtro por tipo de produto
    If objRelABCFornecedoresTela.iTipoProduto <> 0 Then

        'adiciona a tabela Produtos à cláusula From
        sFrom = sFrom & ", Produtos "
        
        'se já existe cláusula where
        If iWhere = 1 Then
            'adiciona cláusula para linkar as tabelas SldDiaForn e Produtos
            sWhere = sWhere & " AND SldDiaForn.Produto = Produtos.Codigo AND Produtos.Tipo =? "
        
        'se não existe
        Else
            'cria cláusula para linkar as tabelas SldDiaForn e Produtos
            sWhere = "WHERE SldDiaForn.Produto = Produtos.Codigo AND Produtos.Tipo =? "
            iWhere = 1
        End If
        
    End If
    '*** TIPODEPRODUTO ***
    
    '*** FORNECEDORESTOP ***
    'Se é para selecionar apenas os xx Produtos mais comprados
    If objRelABCFornecedoresTela.iFornecedorTop > 0 Then
    
        'altera a cláusula select para pegar apenas os TOP xx produtos
        sSelect = Mid(sSelect, 1, 6) & " TOP " & objRelABCFornecedoresTela.iFornecedorTop & Mid(sSelect, 7)
    
    End If
    '*** FORNECEDORESTOP ***
    
    '*** DETALHAMENTO POR FILIAL ***
    'Se é pra obter os valores detalhados por filial de fornecedor
    If objRelABCFornecedoresTela.iDetalharFilial = MARCADO Then
    
        'Inclui o campo filial de fornecedor no select
        sSelect = sSelect & ", SldDiaForn.FilialForn "
        
        'Altera a cláusula GROUP BY, para que o agrupamento seja feito por filial de fornecedor
        sGroupBy = sGroupBy & ", SldDiaForn.FilialForn "
    
    End If
    '*** DETALHAMENTO POR FILIAL ***
    
    '*** CATEGORIA DE PRODUTO ***
    'Verifica se será utilizado filtro por categoria de produto
    If Len(Trim(objRelABCFornecedoresTela.sCategoriaProdutos)) <> 0 Then

        'Altera o select para ler o item de categoria associado ao produto
        'sSelect = sSelect & ", ProdutoCategoria.Item "
        
        'adiciona a tabela Produtos à cláusula From
        sFrom = sFrom & ", ProdutoCategoria "
        
        'se já existe cláusula where
        If iWhere = 1 Then
            
            'adiciona cláusula para linkar as tabelas SldDiaForn e ProdutoCategoria e filtrar pela categoria selecionada
            sWhere = sWhere & " AND SldDiaForn.Produto=ProdutoCategoria.Produto AND ProdutoCategoria.Categoria=? AND (ProdutoCategoria.Item=? "
        
        'se não existe
        Else
            'cria cláusula para linkar as tabelas SldDiaForn e ProdutoCategoria e filtrar pela categoria selecionada
            sWhere = "WHERE SldDiaForn.Produto=ProdutoCategoria.Produto AND ProdutoCategoria.Categoria=? AND (ProdutoCategoria.Item=? "
            iWhere = 1
        
        End If
        
        'Para cada item de categoria selecionado (o loop começa com índice 2, pois o filtro para primeira categoria já foi adicionado acima
        For iIndice = 2 To objRelABCFornecedoresTela.colItensCategoriaProdutos.Count
            'Inclui o filtro na cláusula where
            sWhere = sWhere & "OR ProdutoCategoria.Item=? "
        Next

        'Fecha o parênteses dos itens de categoria na cláusula WHERE
        sWhere = sWhere & ") "

    End If
    '*** CATEGORIA DE PRODUTO ***
    
    '*** CATEGORIA DE FORNECEDOR ***
    'Verifica se será utilizado filtro por categoria de fornecedor
    If Len(Trim(objRelABCFornecedoresTela.sCategoriaFornecedores)) <> 0 Then

        'Altera o select para ler o item de categoria associado à filial do fornecedor
        sSelect = sSelect & ", MAX(FilialFornecedorCategorias.Item) "
        
        'adiciona a tabela Produtos à cláusula From
        sFrom = sFrom & ", FilialFornecedorCategorias "
        
        'se já existe cláusula where
        If iWhere = 1 Then
            'adiciona cláusula para linkar as tabelas SldDiaForn e FilialFornecedorCategorias e filtrar pela categoria selecionada
            sWhere = sWhere & " AND SldDiaForn.Fornecedor=FilialFornecedorCategorias.Fornecedor AND SldDiaForn.FilialForn=FilialFornecedorCategorias.Filial AND FilialFornecedorCategorias.Categoria=? AND (FilialFornecedorCategorias.Item=? "
        
        'se não existe
        Else
            'cria cláusula para linkar as tabelas SldDiaForn e Produtos
            sWhere = "WHERE SldDiaForn.Fornecedor=FilialFornecedorCategorias.Fornecedor AND SldDiaForn.FilialForn=FilialFornecedorCategorias.Filial AND FilialFornecedorCategorias.Categoria=? AND (FilialFornecedorCategorias.Item=? "
            iWhere = 1
        End If
        
        'Para cada item de categoria selecionado
        For iIndice = 2 To objRelABCFornecedoresTela.colItensCategoriaFornecedores.Count
            'Inclui o filtro na cláusula where
            sWhere = sWhere & "OR FilialFornecedorCategorias.Item=? "
        Next

        'Fecha o parênteses dos itens de categoria na cláusula WHERE
        sWhere = sWhere & ") "

    End If
    '*** CATEGORIA DE FORNECEDOR ***
    
    'Monta o select que será executado
    sSelect = sSelect & sFrom & sWhere & sGroupBy & sHaving & sOrderBy
    
    SldDiaForn_Le_RelABCFornecedores1 = SUCESSO

    Exit Function

Erro_SldDiaForn_Le_RelABCFornecedores1:

    SldDiaForn_Le_RelABCFornecedores1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147762)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaForn_Le_RelABCFornecedores2(ByVal lComando As Long, tRelABCFornecedores As TypeRelABCFornecedoresVar, ByVal objRelABCFornecedoresTela As ClassRelABCFornecedoresTela, sSelect As String) As Long
'Faz o bind das variáveis que serão lidas no select

Dim lErro As Long

On Error GoTo Erro_SldDiaForn_Le_RelABCFornecedores2

    With tRelABCFornecedores

        lErro = Comando_PrepararInt(lComando, sSelect)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127042

        .vlFornecedor = CLng(.vlFornecedor)
        lErro = Comando_BindVarInt(lComando, .vlFornecedor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127043

        .vdValor = CDbl(.vdValor)
        lErro = Comando_BindVarInt(lComando, .vdValor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127044
        
        'Se é pra obter os valores detalhados por filial de fornecedor
        If objRelABCFornecedoresTela.iDetalharFilial = MARCADO Then
        
            .viFilialFornecedor = CInt(.viFilialFornecedor)
            lErro = Comando_BindVarInt(lComando, .viFilialFornecedor)
            If (lErro <> AD_SQL_SUCESSO) Then gError 127045
        
        End If
        
        'Se foi passada uma categoria de fornecedor como filtro
        If Len(Trim(objRelABCFornecedoresTela.sCategoriaFornecedores)) > 0 Then
            
            'inicializa a variável que receberá o item de categoria
            .vsItemCategoria = String(STRING_CATEGORIAFORNECEDORITEM_ITEM, 0)
            
            .vsItemCategoria = CStr(.vsItemCategoria)
            lErro = Comando_BindVarInt(lComando, .vsItemCategoria)
            If (lErro <> AD_SQL_SUCESSO) Then gError 127046
        
        End If
        

    End With

    SldDiaForn_Le_RelABCFornecedores2 = SUCESSO

    Exit Function

Erro_SldDiaForn_Le_RelABCFornecedores2:

    SldDiaForn_Le_RelABCFornecedores2 = gErr

    Select Case gErr

        Case 127042 To 127046
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147763)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaForn_Le_RelABCFornecedores3(ByVal lComando As Long, tRelABCFornecedores As TypeRelABCFornecedoresVar, ByVal objRelABCFornecedoresTela As ClassRelABCFornecedoresTela, sSelect As String)
'Faz o bind dos filtros passados para o select e executa o select no BD

Dim lErro As Long
Dim iIndice As Long
Dim sItemCategoriaProduto As String
Dim sItemCategoriaFornecedor As String

On Error GoTo Erro_SldDiaForn_Le_RelABCFornecedores3

    'Se foi passado o filtro DataDe
    If objRelABCFornecedoresTela.dtDataDe <> DATA_NULA Then

        tRelABCFornecedores.vdtDataDe = CDate(objRelABCFornecedoresTela.dtDataDe)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127047

    End If

    'Se foi passado o filtro DataAte
    If objRelABCFornecedoresTela.dtDataAte <> DATA_NULA Then

        tRelABCFornecedores.vdtDataAte = CDate(objRelABCFornecedoresTela.dtDataAte)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127048

    End If

    'Se foi passado filtro FilialEmpresaDe
    If objRelABCFornecedoresTela.iFilialEmpresaDe > 0 Then

        tRelABCFornecedores.viFilialEmpresaDe = CInt(objRelABCFornecedoresTela.iFilialEmpresaDe)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.viFilialEmpresaDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127049

    End If

    'Se foi passado filtro FilialEmpresaAte
    If objRelABCFornecedoresTela.iFilialEmpresaAte > 0 Then

        tRelABCFornecedores.viFilialEmpresaAte = CInt(objRelABCFornecedoresTela.iFilialEmpresaAte)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.viFilialEmpresaAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127050

    End If
    
    'Verifica se foi passado o filtro ProdutoDe
    If Len(Trim(objRelABCFornecedoresTela.sProdutoDe)) > 0 Then

        tRelABCFornecedores.vsProdutoDe = CStr(objRelABCFornecedoresTela.sProdutoDe)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vsProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127051

    End If

    'Verifica se foi passado o filtro ProdutoAte
    If Len(Trim(objRelABCFornecedoresTela.sProdutoAte)) > 0 Then

        tRelABCFornecedores.vsProdutoAte = CStr(objRelABCFornecedoresTela.sProdutoAte)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vsProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127052

    End If
    
    'Verifica se foi passado o filtro FornecedorDe
    If objRelABCFornecedoresTela.lFornecedorDe > 0 Then
    
        tRelABCFornecedores.vlFornecedorDe = CLng(objRelABCFornecedoresTela.lFornecedorDe)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127053

    End If
    
    'Verifica se foi passado o filtro FornecedorAte
    If objRelABCFornecedoresTela.lFornecedorAte > 0 Then
    
        tRelABCFornecedores.vlFornecedorAte = CLng(objRelABCFornecedoresTela.lFornecedorAte)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127054

    End If
    
    'Verifica se foi passado o filtro TipoProdutoDe
    If objRelABCFornecedoresTela.iTipoProduto <> 0 Then

        tRelABCFornecedores.viTipoProduto = CInt(objRelABCFornecedoresTela.iTipoProduto)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.viTipoProduto)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127055

    End If

    'Se foi passado o filtro Categoria de Produto
    If Len(Trim(objRelABCFornecedoresTela.sCategoriaProdutos)) <> 0 Then

        tRelABCFornecedores.vsCategoriaProdutos = CStr(objRelABCFornecedoresTela.sCategoriaProdutos)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vsCategoriaProdutos)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127057

    End If
    
    'Redimensiona o array que armazena os itens de categoria de produtos a serem usados como filtro
    ReDim tRelABCFornecedores.avsItensCategoriaProdutos(objRelABCFornecedoresTela.colItensCategoriaProdutos.Count)
    
    'Para cada item de categoria de produto
    For iIndice = 1 To objRelABCFornecedoresTela.colItensCategoriaProdutos.Count
    
        'Guarda na variável o item de categoria de produto que está na coleção
        sItemCategoriaProduto = objRelABCFornecedoresTela.colItensCategoriaProdutos(iIndice)
        
        'Se o ItemCategoria está preenchido
        If Len(Trim(sItemCategoriaProduto)) <> 0 Then
    
            tRelABCFornecedores.avsItensCategoriaProdutos(iIndice) = CStr(sItemCategoriaProduto)
            lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.avsItensCategoriaProdutos(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then gError 127058
    
        End If
    
    Next
        
    'Se foi passado o filtro Categoria de Fornecedor
    If Len(Trim(objRelABCFornecedoresTela.sCategoriaFornecedores)) <> 0 Then

        tRelABCFornecedores.vsCategoriaFornecedores = CStr(objRelABCFornecedoresTela.sCategoriaFornecedores)
        lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.vsCategoriaFornecedores)
        If (lErro <> AD_SQL_SUCESSO) Then gError 127057

    End If
        
    'Redimensiona o array que armazena os itens de categoria de produtos a serem usados como filtro
    ReDim tRelABCFornecedores.avsItensCategoriaFornecedores(objRelABCFornecedoresTela.colItensCategoriaFornecedores.Count)
    
    'Para cada item de categoria de fornecedor
    For iIndice = 1 To objRelABCFornecedoresTela.colItensCategoriaFornecedores.Count
    
        'Guarda na variável o item de categoria de fornecedor que está na coleção
        sItemCategoriaFornecedor = objRelABCFornecedoresTela.colItensCategoriaFornecedores(iIndice)
        
        'Se o ItemCategoria está preenchido
        If Len(Trim(sItemCategoriaFornecedor)) <> 0 Then
    
            tRelABCFornecedores.avsItensCategoriaFornecedores(iIndice) = CStr(sItemCategoriaFornecedor)
            lErro = Comando_BindVarInt(lComando, tRelABCFornecedores.avsItensCategoriaFornecedores(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then gError 127059
    
        End If
    
    Next
        
    'Executa a leitura no BD
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 127060

    SldDiaForn_Le_RelABCFornecedores3 = SUCESSO

    Exit Function

Erro_SldDiaForn_Le_RelABCFornecedores3:

    SldDiaForn_Le_RelABCFornecedores3 = gErr

    Select Case gErr

        Case 127047 To 127060
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147764)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function SldDiaForn_Le_RelABCFornecedores4(ByVal lComando As Long, tRelABCFornecedores As TypeRelABCFornecedoresVar, ByVal colItensRelABCFornecedores As Collection, ByVal sSelect As String) As Long
'Busca os dados lidos no BD e armazena-os em colItensRelABCFornecedores

Dim lErro As Long
Dim objRelABCFornecedores As ClassRelABCFornecedores
Dim lRanking As Long

On Error GoTo Erro_SldDiaForn_Le_RelABCFornecedores4

    'Busca o primeiro registro lido no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 127061

    'Se não encontrou dados => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 127062

    'Enquanto houver dados no BD
    Do While lErro = AD_SQL_SUCESSO

        'Incrementa a variável que controla a posição do produto no ranking
        lRanking = lRanking + 1

        'Instancia um novo obj
        Set objRelABCFornecedores = New ClassRelABCFornecedores

        'Guarda os dados lidos no obj
        With objRelABCFornecedores
            .lFornecedor = tRelABCFornecedores.vlFornecedor
            .iFilialForn = tRelABCFornecedores.viFilialFornecedor
            .lRanking = lRanking
            .dValor = tRelABCFornecedores.vdValor
            .sItemCategoria = tRelABCFornecedores.vsItemCategoria
        End With

        'Adiciona o obj à coleção
        colItensRelABCFornecedores.Add objRelABCFornecedores

        'Busca o próximo registro lido
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 127063

    Loop

    SldDiaForn_Le_RelABCFornecedores4 = SUCESSO

    Exit Function

Erro_SldDiaForn_Le_RelABCFornecedores4:

    SldDiaForn_Le_RelABCFornecedores4 = gErr

    Select Case gErr

        Case 127061, 127063
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case 127062 'SEM DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147765)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 03/03/04
Private Function RelsABC_Calcula_PercParticipacao(ByVal colItensRelABC As Collection) As Long
'Totaliza o campo valor e calcula o percentual de participação de cada item

Dim objRelABC As Object
Dim dValorTotal As Double

On Error GoTo Erro_RelsABC_Calcula_PercParticipacao

    'Acumula o Valor Total de produtos comprados
    For Each objRelABC In colItensRelABC
        dValorTotal = dValorTotal + CallByName(objRelABC, "dValor", VbGet)
    Next

    'Calcula o percentual do produto de participação de cada produto
    For Each objRelABC In colItensRelABC
        objRelABC.dPercParticipacao = CallByName(objRelABC, "dValor", VbGet) / dValorTotal
    Next
    
    RelsABC_Calcula_PercParticipacao = SUCESSO

    Exit Function

Erro_RelsABC_Calcula_PercParticipacao:

    RelsABC_Calcula_PercParticipacao = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147766)

    End Select
    
    Exit Function

End Function

'###########################################################
'Inserido por Wagner
Function ItemRCItemPC_Le2(ByVal lNumItemPC As Long, ByVal colRC As Collection) As Long
'Lê as RC relacionados ao item de PC passado

Dim lErro As Long
Dim lComando As Long
Dim objRC As ClassRequisicaoCompras
Dim tRC As typeRequisicaoCompras

On Error GoTo Erro_ItemRCItemPC_Le2

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131648
    
    tRC.sOPCodigo = String(STRING_OPCODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RC.NumIntDoc, RC.FilialEmpresa, RC.Codigo, RC.OPCodigo, RC.NumIntDocItemOP FROM ItemRCItemPC AS I, RequisicaoCompraN AS RC, ItensReqCompraN AS IRC  WHERE I.ItemRC = IRC.NumIntDoc AND IRC.ReqCompra = RC.NumIntDoc AND I.ItemPC = ?", _
                                        tRC.lNumIntDoc, tRC.iFilialEmpresa, tRC.lCodigo, tRC.sOPCodigo, tRC.lNumIntDocItemOP, lNumItemPC)
    If lErro <> AD_SQL_SUCESSO Then gError 131649

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131650

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 131651
        
    Do While lErro = AD_SQL_SUCESSO

        Set objRC = New ClassRequisicaoCompras

        objRC.lNumIntDoc = tRC.lNumIntDoc
        objRC.iFilialEmpresa = tRC.iFilialEmpresa
        objRC.lCodigo = tRC.lCodigo
        objRC.sOPCodigo = tRC.sOPCodigo
        objRC.lNumIntDocItemOP = tRC.lNumIntDocItemOP
        
        'Adiciona na coleção
        colRC.Add objRC

        'Busca no BD o próximo item de Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131652

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemRCItemPC_Le2 = SUCESSO

    Exit Function

Erro_ItemRCItemPC_Le2:

    ItemRCItemPC_Le2 = gErr

    Select Case gErr

        Case 131648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131649, 131650, 131652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMPC", gErr)
        
        Case 131651 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147767)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
'#####################################################################

'################################################
'Inserido por Wagner 17/07/2006
Function PCItens_Le_Customizado(ByVal objPC As ClassPedidoCompras) As Long
    PCItens_Le_Customizado = SUCESSO
End Function
'################################################

Function ItemPVItemRC_Le(ByVal lNumItemRC As Long, ByVal colItemPedido As Collection) As Long
'Lê todos os itens de Pedido de Venda relacionados ao item de Requisição de compras passado

Dim lErro As Long
Dim lComando As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntItemPV As Long
Dim dQuantSC As Double

On Error GoTo Erro_ItemPVItemRC_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178645
        
    'Pesquisa os Itens de Pedidos de Venda relacionados a requisicao de compra
    lErro = Comando_Executar(lComando, "SELECT ItemPV, Quantidade FROM ItemPVItemRC WHERE ItemRC = ?", _
        lNumIntItemPV, dQuantSC, lNumItemRC)
    If lErro <> AD_SQL_SUCESSO Then gError 178646

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178647

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPV = New ClassItemPedido

        objItemPV.lNumIntDoc = lNumIntItemPV
        objItemPV.dQuantSC = dQuantSC

        'Adiciona na coleção
        colItemPedido.Add objItemPV

        'Busca no BD o próximo item de PV
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178648

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemPVItemRC_Le = SUCESSO

    Exit Function

Erro_ItemPVItemRC_Le:

    ItemPVItemRC_Le = gErr

    Select Case gErr

        Case 178645
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178646 To 178648
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPVITEMRC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178649)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Cotacao_Le_Req_Associadas(ByVal lNumIntCotacao As Long, ByVal lCodForn As Long, ByVal iFilialForn As Integer, colReq As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim iFilialEmpresa As Integer
Dim objReq As ClassRequisicaoCompras
Dim lCodigoReq As Long
Dim sOPCodigo As String

On Error GoTo Erro_Cotacao_Le_Req_Associadas

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178862
    
    sOPCodigo = String(STRING_OPCODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT RequisicaoCompraN.FilialEmpresa, RequisicaoCompraN.Codigo, MIN(RequisicaoCompraN.OPCodigo) FROM CotacaoProdutoItemRC,CotacaoProdutoN,ItensReqCompraN, RequisicaoCompraN WHERE CotacaoProdutoN.NumIntDoc = CotacaoProdutoItemRC.CotacaoProduto AND CotacaoProdutoN.Cotacao = ? AND CotacaoProdutoItemRC.ItemReqCompra = ItensReqCompraN.NumIntDoc AND ItensReqCompraN.ReqCompra = RequisicaoCompraN.NumIntDoc GROUP BY RequisicaoCompraN.FilialEmpresa, RequisicaoCompraN.Codigo ", _
    iFilialEmpresa, lCodigoReq, sOPCodigo, lNumIntCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 178863
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178864
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objReq = New ClassRequisicaoCompras
    
        objReq.iFilialEmpresa = iFilialEmpresa
        objReq.lCodigo = lCodigoReq
        objReq.sOPCodigo = sOPCodigo
    
        colReq.Add objReq
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178865
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Cotacao_Le_Req_Associadas = SUCESSO
    
    Exit Function

Erro_Cotacao_Le_Req_Associadas:

    Cotacao_Le_Req_Associadas = gErr
    
    Select Case gErr
    
        Case 178862
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178863, 178864, 178865
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178866)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensConc_Le_GeracaoPCOV(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objConcorrencia As ClassConcorrencia) As Long

Dim lErro As Long, sProduto As String, objItemConcorrencia As ClassItemConcorrencia
Dim lComando As Long, lComando2 As Long, sCondPagtoDescRed As String, sFornNomeRed As String, sFilFornNome As String
Dim objCotItemConc As ClassCotacaoItemConc, dQuantComprarMax As Double
Dim tCotItemConc As typeCotacaoItemConcorrencia, sProdutoAnterior As String, sUMCotProd As String
Dim dtData As Date, iCondPagto As Integer, lFornecedor As Long, iFilialForn As Integer, sProdDescricao As String
Dim bAchou As Boolean, lConcCodigo As Long, sConcDescricao As String
Dim colItemConcorrencia As Collection, bSemInfoConcorrencia As Boolean

On Error GoTo Erro_ItensConc_Le_GeracaoPCOV

    bSemInfoConcorrencia = True
    
    Set colItemConcorrencia = objConcorrencia.colItens
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184382
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 184383
    
    sCondPagtoDescRed = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
    sFornNomeRed = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    sFilFornNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)
    tCotItemConc.sMotivoEscolha = String(STRING_MOTIVO_ESCOLHA, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sUMCotProd = String(STRING_UM_SIGLA, 0)
    sProdDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sConcDescricao = String(STRING_NOME, 0)
    
    'obtem as cotacoes atualizadas ligadas 'a concorrencia do orcamento de venda
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ConcorrenciaN.Codigo, ConcorrenciaN.Descricao, CotacaoProdutoN.Produto, ItensCotacaoN.AliquotaIPI, ItensCotacaoN.AliquotaICMS,ItensCotacaoN.CondPagto, ItensCotacaoN.PrecoUnitario, PedidoCotacaoN.Fornecedor, PedidoCotacaoN.Filial, PedidoCotacaoN.Codigo, PedidoCotacaoN.DataValidade, " & _
        " ItensCotacaoN.PrazoEntrega, ItensCotacaoN.QuantEntrega, CotacaoProdutoN.Quantidade, ItensCotacaoN.NumIntDoc, ItensCotacaoN.DataReferencia, PedidoCotacaoN.Data, ItensCotacaoN.Moeda, ItensCotacaoN.Taxa, " & _
        " CondicoesPagto.DescReduzida, Fornecedores.NomeReduzido, FiliaisFornecedores.Nome, CotacaoItemConcorrenciaN.Escolhido, CotacaoItemConcorrenciaN.MotivoEscolha, CotacaoItemConcorrenciaN.PrecoAjustado, CotacaoItemConcorrenciaN.QuantidadeComprar, CotacaoProdutoN.UM, Produtos.Descricao " & _
        " FROM CotacaoProdutoN, ItensCotacaoN, PedidoCotacaoN, ItemPedCotacaoN, CotacaoN, CotacaoProdutoItemFP, ItensFormPreco, ItensOrcamentoVenda, OrcamentoVenda, CondicoesPagto, FiliaisFornecedores, Fornecedores, CotacaoItemConcorrenciaN, ItensConcorrenciaN, ConcorrenciaN, Produtos" & _
        " WHERE Produtos.Codigo = CotacaoProdutoN.Produto AND CotacaoProdutoN.Cotacao = CotacaoN.NumIntDoc AND  CotacaoProdutoN.NumIntDoc = ItemPedCotacaoN.CotacaoProduto AND ItemPedCotacaoN.PedidoCotacao = PedidoCotacaoN.NumIntDoc AND ItensCotacaoN.ItemPedCotacao = ItemPedCotacaoN.NumIntDoc AND " & _
        " CotacaoProdutoItemFP.CotacaoProduto = CotacaoProdutoN.NumIntDoc AND CotacaoProdutoItemFP.NumIntItemFormPreco = ItensFormPreco.NumIntDoc AND ItensFormPreco.NumIntDocOrigem = ItensOrcamentoVenda.NumIntDoc AND ItensFormPreco.TipoDocOrigem = ? " & _
        " AND ItensOrcamentoVenda.FilialEmpresa = OrcamentoVenda.FilialEmpresa AND ItensOrcamentoVenda.CodOrcamento = OrcamentoVenda.Codigo AND OrcamentoVenda.FilialEmpresa = ? AND OrcamentoVenda.Codigo = ?" & _
        " AND ItensCotacaoN.CondPagto = CondicoesPagto.Codigo AND FiliaisFornecedores.CodFornecedor = Fornecedores.Codigo AND PedidoCotacaoN.Fornecedor = FiliaisFornecedores.CodFornecedor AND PedidoCotacaoN.Filial = FiliaisFornecedores.CodFilial" & _
        " AND CotacaoItemConcorrenciaN.ItemCotacao = ItensCotacaoN.NumIntDoc AND CotacaoItemConcorrenciaN.ItemConcorrencia = ItensConcorrenciaN.NumIntDoc AND ItensConcorrenciaN.Concorrencia = ConcorrenciaN.NumIntDoc" & _
        " ORDER BY CotacaoProdutoN.Produto, PedidoCotacaoN.Fornecedor, PedidoCotacaoN.Filial, PedidoCotacaoN.Data DESC, ItensCotacaoN.CondPagto, PedidoCotacaoN.Codigo DESC, ItensCotacaoN.PrecoUnitario", _
        lConcCodigo, sConcDescricao, sProduto, tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
        tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, _
        sCondPagtoDescRed, sFornNomeRed, sFilFornNome, tCotItemConc.iEscolhido, tCotItemConc.sMotivoEscolha, tCotItemConc.dPrecoAjustado, tCotItemConc.dQuantidadeComprar, sUMCotProd, sProdDescricao, FPORIGEM_ITEMOV, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 184384
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184385
    
    sProdutoAnterior = ""
    
    Do While lErro = AD_SQL_SUCESSO
    
        If bSemInfoConcorrencia Then
            If lConcCodigo <> 0 Then
                objConcorrencia.lCodigo = lConcCodigo
                objConcorrencia.sDescricao = sConcDescricao
                bSemInfoConcorrencia = False
            End If
        Else
            If lConcCodigo <> 0 And lConcCodigo <> objConcorrencia.lCodigo Then gError 184408
        End If
        
        If sProdutoAnterior <> sProduto Then
            
            Set objItemConcorrencia = New ClassItemConcorrencia
            
            objItemConcorrencia.sProduto = sProduto
            objItemConcorrencia.sDescricao = sProdDescricao
            objItemConcorrencia.sUM = sUMCotProd
            objItemConcorrencia.dtDataNecessidade = DATA_NULA
            objItemConcorrencia.dQuantidade = tCotItemConc.dQuantidadeComprar
            
            colItemConcorrencia.Add objItemConcorrencia
            sProdutoAnterior = sProduto
            
        End If
        
        Set objCotItemConc = New ClassCotacaoItemConc
        
        objCotItemConc.sCondPagto = iCondPagto & SEPARADOR & sCondPagtoDescRed
        objCotItemConc.lItemCotacao = tCotItemConc.lItemCotacao
        objCotItemConc.dPrecoUnitario = tCotItemConc.dPrecoUnitario
        objCotItemConc.dtDataPedidoCotacao = dtData
        objCotItemConc.sFornecedor = sFornNomeRed
        objCotItemConc.sFilial = iFilialForn & SEPARADOR & sFilFornNome
        
        objCotItemConc.lPedCotacao = tCotItemConc.lPedCotacao
        objCotItemConc.dtDataValidade = tCotItemConc.dtDataValidade
        objCotItemConc.iPrazoEntrega = tCotItemConc.iPrazoEntrega
        objCotItemConc.dQuantEntrega = tCotItemConc.dQuantEntrega
        objCotItemConc.dAliquotaIPI = tCotItemConc.dAliquotaIPI
        objCotItemConc.dAliquotaICMS = tCotItemConc.dCreditoICMS
                
        objCotItemConc.sUMCompra = sUMCotProd
        objCotItemConc.dQuantCotada = dQuantComprarMax
        objCotItemConc.dPrecoAjustado = objCotItemConc.dPrecoUnitario
        objCotItemConc.iSelecionada = MARCADO
        objCotItemConc.iMoeda = tCotItemConc.iMoeda
        objCotItemConc.dTaxa = tCotItemConc.dTaxa
        If objCotItemConc.dTaxa = 0 Then
        End If
        If tCotItemConc.dtDataEntrega <> DATA_NULA Then
            objCotItemConc.dtDataEntrega = DateAdd("d", tCotItemConc.iPrazoEntrega, tCotItemConc.dtDataEntrega)
        Else
            objCotItemConc.dtDataEntrega = DATA_NULA
        End If
        
        'já gravados na concorrencia
        objCotItemConc.dQuantidadeComprar = tCotItemConc.dQuantidadeComprar
        objCotItemConc.dPrecoAjustado = tCotItemConc.dPrecoAjustado
        objCotItemConc.iEscolhido = tCotItemConc.iEscolhido
        objCotItemConc.sMotivoEscolha = tCotItemConc.sMotivoEscolha
        
        'Se estiver Adiciona na coleção
        objItemConcorrencia.colCotacaoItemConc.Add objCotItemConc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184386
    
    Loop
    
    sCondPagtoDescRed = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
    sFornNomeRed = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    sFilFornNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sUMCotProd = String(STRING_UM_SIGLA, 0)
    sProdDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    'obtem as cotecoes atualizadas ligadas ao orcamento de vendas,
    'estando ou não ligadas à concorrencia
    lErro = Comando_Executar(lComando2, "SELECT DISTINCT CotacaoProdutoN.Produto, ItensCotacaoN.AliquotaIPI, ItensCotacaoN.AliquotaICMS,ItensCotacaoN.CondPagto, ItensCotacaoN.PrecoUnitario, PedidoCotacaoN.Fornecedor, PedidoCotacaoN.Filial, PedidoCotacaoN.Codigo, PedidoCotacaoN.DataValidade, " & _
        " ItensCotacaoN.PrazoEntrega, ItensCotacaoN.QuantEntrega, CotacaoProdutoN.Quantidade, ItensCotacaoN.NumIntDoc, ItensCotacaoN.DataReferencia, PedidoCotacaoN.Data, ItensCotacaoN.Moeda, ItensCotacaoN.Taxa, " & _
        " CondicoesPagto.DescReduzida, Fornecedores.NomeReduzido, FiliaisFornecedores.Nome, CotacaoProdutoN.UM, Produtos.Descricao " & _
        " FROM CotacaoProdutoN, ItensCotacaoN, PedidoCotacaoN, ItemPedCotacaoN, CotacaoN, CotacaoProdutoItemFP, ItensFormPreco, ItensOrcamentoVenda, OrcamentoVenda, CondicoesPagto, FiliaisFornecedores, Fornecedores, Produtos" & _
        " WHERE Produtos.Codigo = CotacaoProdutoN.Produto AND CotacaoProdutoN.Cotacao = CotacaoN.NumIntDoc AND  CotacaoProdutoN.NumIntDoc = ItemPedCotacaoN.CotacaoProduto AND ItemPedCotacaoN.PedidoCotacao = PedidoCotacaoN.NumIntDoc AND ItensCotacaoN.ItemPedCotacao = ItemPedCotacaoN.NumIntDoc AND " & _
        " CotacaoProdutoItemFP.CotacaoProduto = CotacaoProdutoN.NumIntDoc AND CotacaoProdutoItemFP.NumIntItemFormPreco = ItensFormPreco.NumIntDoc AND ItensFormPreco.NumIntDocOrigem = ItensOrcamentoVenda.NumIntDoc AND ItensFormPreco.TipoDocOrigem = ? " & _
        " AND ItensOrcamentoVenda.FilialEmpresa = OrcamentoVenda.FilialEmpresa AND ItensOrcamentoVenda.CodOrcamento = OrcamentoVenda.Codigo AND OrcamentoVenda.FilialEmpresa = ? AND OrcamentoVenda.Codigo = ?" & _
        " AND ItensCotacaoN.CondPagto = CondicoesPagto.Codigo AND FiliaisFornecedores.CodFornecedor = Fornecedores.Codigo AND PedidoCotacaoN.Fornecedor = FiliaisFornecedores.CodFornecedor AND PedidoCotacaoN.Filial = FiliaisFornecedores.CodFilial" & _
        " ORDER BY CotacaoProdutoN.Produto, PedidoCotacaoN.Fornecedor, PedidoCotacaoN.Filial, PedidoCotacaoN.Data DESC, ItensCotacaoN.CondPagto, PedidoCotacaoN.Codigo DESC, ItensCotacaoN.PrecoUnitario", _
        sProduto, tCotItemConc.dAliquotaIPI, tCotItemConc.dCreditoICMS, iCondPagto, tCotItemConc.dPrecoUnitario, lFornecedor, iFilialForn, tCotItemConc.lPedCotacao, tCotItemConc.dtDataValidade, tCotItemConc.iPrazoEntrega, _
        tCotItemConc.dQuantEntrega, dQuantComprarMax, tCotItemConc.lItemCotacao, tCotItemConc.dtDataEntrega, dtData, tCotItemConc.iMoeda, tCotItemConc.dTaxa, _
        sCondPagtoDescRed, sFornNomeRed, sFilFornNome, sUMCotProd, sProdDescricao, FPORIGEM_ITEMOV, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 184387
    
    lErro = Comando_BuscarProximo(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184388
    
    sProdutoAnterior = ""
    
    Do While lErro = AD_SQL_SUCESSO
    
        If sProdutoAnterior <> sProduto Then
            
            bAchou = False
            For Each objItemConcorrencia In colItemConcorrencia
            
                If objItemConcorrencia.sProduto = sProduto Then
                    bAchou = True
                    Exit For
                End If
                
            Next
            
            If bAchou = False Then
            
                Set objItemConcorrencia = New ClassItemConcorrencia
                objItemConcorrencia.sProduto = sProduto
                objItemConcorrencia.sDescricao = sProdDescricao
                objItemConcorrencia.sUM = sUMCotProd
                objItemConcorrencia.dtDataNecessidade = DATA_NULA
                objItemConcorrencia.dQuantidade = dQuantComprarMax
                
                colItemConcorrencia.Add objItemConcorrencia
            
            End If
            
            sProdutoAnterior = sProduto
            
        End If
        
        bAchou = False
        For Each objCotItemConc In objItemConcorrencia.colCotacaoItemConc
        
            If objCotItemConc.lItemCotacao = tCotItemConc.lItemCotacao Then
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou = False Then
        
            Set objCotItemConc = New ClassCotacaoItemConc
            
            objCotItemConc.sCondPagto = iCondPagto & SEPARADOR & sCondPagtoDescRed
            objCotItemConc.lItemCotacao = tCotItemConc.lItemCotacao
            objCotItemConc.dPrecoUnitario = tCotItemConc.dPrecoUnitario
            objCotItemConc.dtDataPedidoCotacao = dtData
            objCotItemConc.sFornecedor = sFornNomeRed
            objCotItemConc.sFilial = iFilialForn & SEPARADOR & sFilFornNome
            
            objCotItemConc.lPedCotacao = tCotItemConc.lPedCotacao
            objCotItemConc.dtDataValidade = tCotItemConc.dtDataValidade
            objCotItemConc.iPrazoEntrega = tCotItemConc.iPrazoEntrega
            objCotItemConc.dQuantEntrega = tCotItemConc.dQuantEntrega
            objCotItemConc.dAliquotaIPI = tCotItemConc.dAliquotaIPI
            objCotItemConc.dAliquotaICMS = tCotItemConc.dCreditoICMS
                    
            objCotItemConc.sUMCompra = sUMCotProd
            objCotItemConc.dQuantCotada = dQuantComprarMax
            objCotItemConc.dPrecoAjustado = objCotItemConc.dPrecoUnitario
            objCotItemConc.iSelecionada = MARCADO
            objCotItemConc.iMoeda = tCotItemConc.iMoeda
            objCotItemConc.dTaxa = tCotItemConc.dTaxa
            If tCotItemConc.dtDataEntrega <> DATA_NULA Then
                objCotItemConc.dtDataEntrega = DateAdd("d", tCotItemConc.iPrazoEntrega, tCotItemConc.dtDataEntrega)
            Else
                objCotItemConc.dtDataEntrega = DATA_NULA
            End If
            
            'atributos gravados com a concorrencia ainda nao estao disponiveis
            objCotItemConc.dQuantidadeComprar = dQuantComprarMax
            objCotItemConc.dPrecoAjustado = objCotItemConc.dPrecoUnitario
            objCotItemConc.iEscolhido = 0
            objCotItemConc.sMotivoEscolha = ""
            
            'Se estiver Adiciona na coleção
            objItemConcorrencia.colCotacaoItemConc.Add objCotItemConc
        
        End If
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184389
    
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    ItensConc_Le_GeracaoPCOV = SUCESSO
    
    Exit Function
    
Erro_ItensConc_Le_GeracaoPCOV:

    ItensConc_Le_GeracaoPCOV = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184380)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function Produto_Le_Quantidades_Compras(objProduto As ClassProduto, dQuantPedidaReqCompra As Double, dQuantidadePedCompra As Double, ByVal iFilialEmpresa As Integer, Optional ByVal bExibeNaoEnviados As Boolean = False) As Long
'Realiza a leitura da QuantidadePedida na Requisição e a Quantidade no Pedido de Compras

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim dFator As Double
Dim sUM As String, dtDataFiltro As Date

On Error GoTo Erro_Produto_Le_Quantidades_Compras

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125320
    Next

    dQuantPedidaReqCompra = 0
    dQuantidadePedCompra = 0
    
    If bExibeNaoEnviados Then
        dtDataFiltro = #1/1/1822# 'Apenas para trazer todo mundo e não ter que mudar muito o SELECT
    Else
        dtDataFiltro = DATA_NULA
    End If

    'Inicializa as strings
    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
  
    'Busca no BD o código passado
    lErro = Comando_Executar(alComando(1), "SELECT ItensReqCompra.UM, SUM(ItensReqCompra.Quantidade - ItensReqCompra.QuantPedida - ItensReqCompra.QuantRecebida - ItensReqCompra.QuantCancelada) FROM Produtos, ProdutosFilial, ItensReqCompra, RequisicaoCompra WHERE Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.Produto = ItensReqCompra.Produto AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.DataEnvio <> ? AND Produtos.Codigo = ? AND ProdutosFilial.FilialEmpresa = ? AND (ItensReqCompra.Quantidade - ItensReqCompra.QuantPedida - ItensReqCompra.QuantRecebida - ItensReqCompra.QuantCancelada) > 0 GROUP BY ItensReqCompra.Produto, ItensReqCompra.UM", sUM, dQuantidade, dtDataFiltro, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 125321
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125322
    
    'Se nao encontrou ==> erro
'    If lErro <> AD_SQL_SUCESSO Then gError 125323
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Realiza a converção para um só UM
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 125324
        
        'Preenche a variável quantidade
        dQuantPedidaReqCompra = dQuantPedidaReqCompra + (dQuantidade * dFator)
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125325
        
    Loop
    
    'Inicializa as strings
    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
    'Busca no BD o pedido com o produto passado
    lErro = Comando_Executar(alComando(2), "SELECT ItensPedCompra.UM, SUM(ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento) FROM Produtos, ProdutosFilial, ItensPedCompra, PedidoCompra WHERE Produtos.Codigo = ProdutosFilial.Produto And ProdutosFilial.Produto = ItensPedCompra.Produto AND ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND PedidoCompra.DataEnvio <> ? AND Produtos.Codigo = ? AND ProdutosFilial.FilialEmpresa = ? AND (ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento) > 0 GROUP BY ItensPedCompra.Produto, ItensPedCompra.UM", sUM, dQuantidade, dtDataFiltro, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 125326
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125327
    
    'Se nao encontrou ==> erro
'    If lErro <> AD_SQL_SUCESSO Then gError 125328
    
    Do While lErro = AD_SQL_SUCESSO
        
        'Realiza a converção para uma mesma UM
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 125329
            
        'Preenche a quantidade
        dQuantidadePedCompra = dQuantidadePedCompra + (dQuantidade * dFator)
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125330
        
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Le_Quantidades_Compras = SUCESSO
    
    Exit Function
    
Erro_Produto_Le_Quantidades_Compras:
    
    Produto_Le_Quantidades_Compras = gErr
    
    Select Case gErr
    
        Case 125320
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 125321, 125322, 125325, 125326, 125327, 125330
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objProduto.sCodigo)
        
        Case 125323, 125328 'Não encotrou o elemento
        
        Case 125324, 125329
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155813)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function ReqComprasEnvio_Le(ByVal objReqEnvio As ClassReqCompraEnvio) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim sSQL As String
Dim lNumIntDoc As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras, objProdFil As ClassProdutoFilial
Dim objProd As ClassProduto

On Error GoTo Erro_ReqComprasEnvio_Le

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211106
    Next
    
    Set objReqEnvio.colRequisicao = New Collection
    
    lErro = ReqComprasEnvio_PreparaSQL(objReqEnvio, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = ReqComprasEnvio_PreparaInt(alComando(1), sSQL, objReqEnvio.dtDataDe, objReqEnvio.dtDataAte, objReqEnvio.dtDataLimiteDe, objReqEnvio.dtDataLimiteAte, objReqEnvio.dtDataLimCDe, objReqEnvio.dtDataLimCAte, objReqEnvio.lCodigoDe, objReqEnvio.lCodigoAte, lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211107

    'Enquanto existirem dados de acordo com o filtro passado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objReqCompra = New ClassRequisicaoCompras
    
        objReqCompra.lNumIntDoc = lNumIntDoc
    
        lErro = CF("RequisicaoCompras_Le", objReqCompra)
        If lErro <> SUCESSO And lErro <> 54397 Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("ItensReqCompras_Le", objReqCompra)
        If lErro <> SUCESSO And lErro <> 54377 Then gError ERRO_SEM_MENSAGEM
        
        For Each objItemRC In objReqCompra.colItens
        
            Set objProdFil = New ClassProdutoFilial
            Set objProd = New ClassProduto
            
            objProdFil.sProduto = objItemRC.sProduto
            objProdFil.iFilialEmpresa = objReqCompra.iFilialEmpresa
        
            lErro = CF("ProdutoFilial_Le", objProdFil)
            If lErro <> SUCESSO And lErro <> 28261 Then gError ERRO_SEM_MENSAGEM
            
            objProd.sCodigo = objItemRC.sProduto
            
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            objItemRC.sDescProduto = objProd.sDescricao
            objItemRC.iTempoRessup = objProdFil.iTempoRessup
        
        Next
        
        objReqEnvio.colRequisicao.Add objReqCompra
            
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211108
    
    Loop
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ReqComprasEnvio_Le = SUCESSO
     
    Exit Function
    
Erro_ReqComprasEnvio_Le:

    ReqComprasEnvio_Le = gErr
     
    Select Case gErr
          
        Case 211106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211107, 211108
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211109)
     
    End Select
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function ReqComprasEnvio_PreparaSQL(ByVal objReqEnvio As ClassReqCompraEnvio, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sOrderBy As String

On Error GoTo Erro_ReqComprasEnvio_PreparaSQL

    sSQL = "SELECT R.NumIntDoc FROM RequisicaoCompra AS R "
    sWhere = "WHERE R.FilialEmpresa = ? AND " & IIf(objReqEnvio.iTipo = 2, "R.DataEnvio <> ? AND R.DataRegAprov = ?", "R.DataEnvio = ? AND R.DataEnvio = ?") & " AND R.DataBaixa = ? "
    sOrderBy = "ORDER BY R.Codigo "
    
    If objReqEnvio.lCodigoDe <> 0 Then
        sWhere = sWhere & "AND R.Codigo >= ? "
    End If
    
    If objReqEnvio.lCodigoAte <> 0 Then
        sWhere = sWhere & "AND R.Codigo <= ? "
    End If
    
    If objReqEnvio.dtDataDe <> DATA_NULA Then
        sWhere = sWhere & "AND R.Data >= ? "
    End If
    
    If objReqEnvio.dtDataAte <> DATA_NULA Then
        sWhere = sWhere & "AND R.Data <= ? "
    End If
    
    If objReqEnvio.dtDataLimiteDe <> DATA_NULA Then
        sWhere = sWhere & "AND R.DataLimite >= ? "
    End If
    
    If objReqEnvio.dtDataLimiteAte <> DATA_NULA Then
        sWhere = sWhere & "AND R.DataLimite <= ? "
    End If
    
    If objReqEnvio.dtDataLimCDe <> DATA_NULA Or objReqEnvio.dtDataLimCAte <> DATA_NULA Then
    
        sWhere = sWhere & "AND R.NumIntDoc IN (SELECT Z.NumIntDoc FROM ItensReqCompraN AS I, RequisicaoCompraN AS Z, ProdutosFilial AS P " & _
                        "WHERE I.ReqCompra = Z.NumIntDoc AND I.Produto = P.Produto AND Z.DataLimite <> ? "
        
        If objReqEnvio.dtDataLimCDe <> DATA_NULA Then
            sWhere = sWhere & "AND DateAdd(Day, -P.TempoRessup, Z.DataLimite) >= ? "
        End If
        
        If objReqEnvio.dtDataLimCAte <> DATA_NULA Then
            sWhere = sWhere & "AND DateAdd(Day, -P.TempoRessup, Z.DataLimite) <= ? "
        End If
        
        sWhere = sWhere & ")"
        
    End If
    
    sSQL = sSQL & sWhere & sOrderBy
    
    ReqComprasEnvio_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_ReqComprasEnvio_PreparaSQL:

    ReqComprasEnvio_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211110)
     
    End Select
     
    Exit Function

End Function

Private Function ReqComprasEnvio_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal vdtDataDe As Variant, ByVal vdtDataAte As Variant, ByVal vdtDataLimiteDe As Variant, ByVal vdtDataLimiteAte As Variant, ByVal vdtDataLimCDe As Variant, ByVal vdtDataLimCAte As Variant, ByVal vlCodigoDe As Variant, ByVal vlCodigoAte As Variant, vlNumIntDoc As Variant) As Long

Dim lErro As Long
Dim vdtDataNula As Variant
Dim viFilialEmpresa As Variant

On Error GoTo Erro_ReqComprasEnvio_PreparaInt

    vdtDataNula = DATA_NULA
    viFilialEmpresa = giFilialEmpresa

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211111

    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211112
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119

    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    If vlCodigoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211113
    End If
    
    If vlCodigoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211114
    End If
    
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211115
    End If
    
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211116
    End If
    
    If vdtDataLimiteDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211117
    End If
    
    If vdtDataLimiteAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataLimiteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211118
    End If
    
    If vdtDataLimCDe <> DATA_NULA Or vdtDataLimCAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataNula)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    End If
    
    If vdtDataLimCDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataLimCDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211120
    End If
    
    If vdtDataLimCAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataLimCAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211121
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211122
    
    ReqComprasEnvio_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_ReqComprasEnvio_PreparaInt:

    ReqComprasEnvio_PreparaInt = gErr
     
    Select Case gErr

        Case 211111 To 211122
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211123)
     
    End Select
     
    Exit Function

End Function

Function ItensPC_Le_Produto(ByVal iFilialEmpresa As Integer, ByVal sCodProduto As String, colItensPC As Collection) As Long
'Lê no BD a coleção de ítens em aberto de um determinado produto com qtde a receber e devolve em colItensPC

Dim lErro As Long
Dim lComando1 As Long, lComando2 As Long
Dim objItemPC As ClassItemPedCompra
Dim lCodigoPC As Long, dtDataLimite As Date
Dim dQuantidade As Double, lNumIntDoc As Long
Dim dtDataEntrega As Date, dQtdeEntrega As Double, dQtdeEntregaPC As Double
Dim dSobra As Double, sUM As String

On Error GoTo Erro_ItensPC_Le_Produto

    'Abre os comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 210762

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 210762
    
    sUM = String(STRING_UM_SIGLA, 0)

    'Le da tabela de ItensPedCompra a colecao de itens do Pedido de Compra passado como parametro
    lErro = Comando_Executar(lComando1, "SELECT I.NumIntDoc, I.UM, Codigo, Quantidade, Quantidade - QuantRecebida - QuantRecebimento, DataLimite FROM ItensPedCompra AS I, PedidoCompra AS P WHERE P.NumIntDoc = I.PedCompra AND Quantidade - QuantRecebida - QuantRecebimento > 0 AND P.FilialEmpresa  = ? AND I.Produto = ? ORDER BY Codigo", lNumIntDoc, sUM, lCodigoPC, dQtdeEntregaPC, dQuantidade, dtDataLimite, iFilialEmpresa, sCodProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 210763

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210764

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_Executar(lComando2, "SELECT DataEntrega, Quantidade FROM ItensPCEntrega WHERE NumIntItemPC = ? ORDER BY DataEntrega", dtDataEntrega, dQtdeEntrega, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 210763
    
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210764
        
        If lErro = AD_SQL_SEM_DADOS Then

            Set objItemPC = New ClassItemPedCompra
    
            'preenchendo objitem
            objItemPC.lCodigoPC = lCodigoPC
            objItemPC.dQuantidade = dQuantidade
            objItemPC.dtDataLimite = dtDataLimite
            objItemPC.sUM = sUM
            
            'Insere em colItens os valores de objItemPC
            colItensPC.Add objItemPC
            
        Else
        
            dSobra = dQtdeEntregaPC - dQuantidade
        
            Do While lErro = AD_SQL_SUCESSO
        
                Set objItemPC = New ClassItemPedCompra
        
                'preenchendo objitem
                objItemPC.lCodigoPC = lCodigoPC
                objItemPC.sUM = sUM
                
                If dSobra > QTDE_ESTOQUE_DELTA Then
                
                    If dSobra > dQtdeEntrega Then
                        objItemPC.dQuantidade = 0
                        dSobra = dSobra - dQtdeEntrega
                    Else
                        objItemPC.dQuantidade = dQtdeEntrega - dSobra
                        dSobra = 0
                    End If
                
                Else
                    objItemPC.dQuantidade = dQtdeEntrega
                End If
                
                objItemPC.dtDataLimite = dtDataEntrega
                
                'Insere em colItens os valores de objItemPC
                If objItemPC.dQuantidade > QTDE_ESTOQUE_DELTA Then colItensPC.Add objItemPC
        
                lErro = Comando_BuscarProximo(lComando2)
                If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210765
        
            Loop
        
        End If

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210765
        
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    ItensPC_Le_Produto = SUCESSO

    Exit Function

Erro_ItensPC_Le_Produto:

    ItensPC_Le_Produto = gErr

    Select Case gErr

        Case 210762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 210763, 210764, 210765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210766)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function PedidoCompra_Valida_Campos_Cust(ByVal objPedidoCompra As ClassPedidoCompras) As Long
    PedidoCompra_Valida_Campos_Cust = SUCESSO
End Function

Public Function PedComprasEnvio_Le(ByVal objPedEnvio As ClassPedCompraEnvio) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim sSQL As String
Dim lNumIntDoc As Long
Dim objPedCompra As ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim objProd As ClassProduto

On Error GoTo Erro_PedComprasEnvio_Le

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211106
    Next
    
    Set objPedEnvio.colPedidos = New Collection
    
    lErro = PedComprasEnvio_PreparaSQL(objPedEnvio, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = PedComprasEnvio_PreparaInt(alComando(1), sSQL, objPedEnvio.dtDataDe, objPedEnvio.dtDataAte, objPedEnvio.lCodigoDe, objPedEnvio.lCodigoAte, lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211107

    'Enquanto existirem dados de acordo com o filtro passado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPedCompra = New ClassPedidoCompras
    
        objPedCompra.lNumIntDoc = lNumIntDoc
    
        lErro = CF("PedidoCompra_Le_Todos", objPedCompra)
        If lErro <> SUCESSO And lErro <> 68486 Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("ItensPC_LeTodos", objPedCompra)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For Each objItemPC In objPedCompra.colItens
        
            Set objProd = New ClassProduto

            objProd.sCodigo = objItemPC.sProduto
            
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            objItemPC.sDescProduto = objProd.sDescricao
        
        Next
        
        objPedEnvio.colPedidos.Add objPedCompra
            
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211108
    
    Loop
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedComprasEnvio_Le = SUCESSO
     
    Exit Function
    
Erro_PedComprasEnvio_Le:

    PedComprasEnvio_Le = gErr
     
    Select Case gErr
          
        Case 211106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211107, 211108
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211109)
     
    End Select
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function PedComprasEnvio_PreparaSQL(ByVal objPedEnvio As ClassPedCompraEnvio, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sOrderBy As String

On Error GoTo Erro_PedComprasEnvio_PreparaSQL

    sSQL = "SELECT P.NumIntDoc FROM PedidoCompra AS P "
    sWhere = "WHERE P.FilialEmpresa = ? AND P.DataEnvio = ? AND P.DataBaixa = ? AND P.DataRegAprov = ? "
    sOrderBy = "ORDER BY P.Codigo "
    
    If objPedEnvio.lCodigoDe <> 0 Then
        sWhere = sWhere & "AND P.Codigo >= ? "
    End If
    
    If objPedEnvio.lCodigoAte <> 0 Then
        sWhere = sWhere & "AND P.Codigo <= ? "
    End If
    
    If objPedEnvio.dtDataDe <> DATA_NULA Then
        sWhere = sWhere & "AND P.Data >= ? "
    End If
    
    If objPedEnvio.dtDataAte <> DATA_NULA Then
        sWhere = sWhere & "AND P.Data <= ? "
    End If
    
    sSQL = sSQL & sWhere & sOrderBy
    
    PedComprasEnvio_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_PedComprasEnvio_PreparaSQL:

    PedComprasEnvio_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211110)
     
    End Select
     
    Exit Function

End Function

Private Function PedComprasEnvio_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal vdtDataDe As Variant, ByVal vdtDataAte As Variant, ByVal vlCodigoDe As Variant, ByVal vlCodigoAte As Variant, vlNumIntDoc As Variant) As Long

Dim lErro As Long
Dim vdtDataNula As Variant
Dim viFilialEmpresa As Variant

On Error GoTo Erro_PedComprasEnvio_PreparaInt

    vdtDataNula = DATA_NULA
    viFilialEmpresa = giFilialEmpresa

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211111

    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211112
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211119
    
    If vlCodigoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211113
    End If
    
    If vlCodigoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211114
    End If
    
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211115
    End If
    
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211116
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211122
    
    PedComprasEnvio_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_PedComprasEnvio_PreparaInt:

    PedComprasEnvio_PreparaInt = gErr
     
    Select Case gErr

        Case 211111 To 211122
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211123)
     
    End Select
     
    Exit Function

End Function

Function ItensPCEntrega_Le(ByVal objPC As ClassPedidoCompras) As Long

Dim lComando As Long
Dim objDataEntrega As ClassDataEntrega
Dim objItemPC As ClassItemPedCompra
Dim lErro As Long
Dim lNumIntPC As Long, dtDataEntrega As Date, dQuantidade As Double

On Error GoTo Erro_ItensPCEntrega_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213204
    
    For Each objItemPC In objPC.colItens
               
        lErro = Comando_Executar(lComando, "SELECT NumIntPC, DataEntrega, Quantidade FROM ItensPCEntrega WHERE NumIntItemPC = ?", _
        lNumIntPC, dtDataEntrega, dQuantidade, objItemPC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 213205
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213206
    
        Do While lErro <> AD_SQL_SEM_DADOS
            
            Set objDataEntrega = New ClassDataEntrega
    
            With objDataEntrega
                            
                .dtDataEntrega = dtDataEntrega
                .dQuantidade = dQuantidade
            
            End With
            
            objItemPC.colDataEntrega.Add objDataEntrega
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213207
        
        Loop
    
    Next
    
    Call Comando_Fechar(lComando)
            
    ItensPCEntrega_Le = SUCESSO

    Exit Function

Erro_ItensPCEntrega_Le:

    ItensPCEntrega_Le = gErr

    Select Case gErr

        Case 213204
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 213205 To 213207
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPCENTREGA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213208)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

