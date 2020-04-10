VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRBSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ICMSAliqExternas_Le_EstadoOrigem(ByVal sSiglaOrigem As String, ByVal colAliqExternas As Collection) As Long
'carrega as aliquotas externas que partem de um estado (sSiglaOrigem)

Dim lErro As Long, lComando As Long, lComando2 As Long, sSiglaDest As String, dAliquota As Double
Dim objICMSAliqExterna As ClassICMSAliqExterna, dAliquotaFCP As Double

On Error GoTo Erro_ICMSAliqExternas_Le_EstadoOrigem

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27268
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 27268
    
    sSiglaDest = String(STRING_ESTADO, 0)
    
    'ler ICMSAliquotasExternas
    lErro = Comando_Executar(lComando, "SELECT SiglaDest, Aliquota, AliquotaFCP FROM ICMSAliquotasExternas WHERE SiglaOrig = ?", sSiglaDest, dAliquota, dAliquotaFCP, sSiglaOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 27269
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27270
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objICMSAliqExterna = New ClassICMSAliqExterna
        
        objICMSAliqExterna.dAliquota = dAliquota
        objICMSAliqExterna.sSiglaDest = sSiglaDest
        objICMSAliqExterna.sSiglaOrig = sSiglaOrigem
        
'        lErro = Comando_Executar(lComando2, "SELECT ICMSPercFCP FROM Estados WHERE Sigla = ?", dAliquotaFCP, sSiglaOrigem)
'        If lErro <> AD_SQL_SUCESSO Then Error 27269
'
'        lErro = Comando_BuscarProximo(lComando2)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27270
'
        objICMSAliqExterna.dAliquotaFCP = dAliquotaFCP
        
        colAliqExternas.Add objICMSAliqExterna, sSiglaDest
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27271
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)
    
    ICMSAliqExternas_Le_EstadoOrigem = SUCESSO

    Exit Function

Erro_ICMSAliqExternas_Le_EstadoOrigem:

    ICMSAliqExternas_Le_EstadoOrigem = Err

    Select Case Err

        Case 27268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27269, 27270, 27271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_ALIQ_EXT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153657)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function ICMSExcecoes_Le_Estado(ByVal sEstado As String, ByVal colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassICMSExcecoes de um determinado estado, ordenandos pelo campo "prioridade"
Dim lErro As Long, lComando As Long
Dim objICMSExcecao As ClassICMSExcecao, tICMSExcecao As typeICMSExcecao
On Error GoTo Erro_ICMSExcecoes_Le_Estado

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27272
    
    tICMSExcecao.sCategoriaCliente = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    tICMSExcecao.sCategoriaClienteItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    tICMSExcecao.sCategoriaProduto = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tICMSExcecao.sCategoriaProdutoItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    '###########################################################
    'Inserido por Wagner 29/09/05
    tICMSExcecao.sCategoriaFornecedor = String(STRING_CATEGORIAFORNECEDOR_CATEGORIA, 0)
    tICMSExcecao.sCategoriaFornecedorItem = String(STRING_CATEGORIAFORNECEDORITEM_ITEM, 0)
    '###########################################################
    
    tICMSExcecao.sEstadoOrigem = String(STRING_ESTADO_SIGLA, 0)
    tICMSExcecao.scBenef = String(20, 0)
    '###########################################################
    'Alterado por Wagner 29/09/05
    lErro = Comando_Executar(lComando, "SELECT GrupoOrigemMercadoria, EstadoOrigem, CategoriaProduto, CategoriaProdutoItem, CategoriaCliente, CategoriaClienteItem, Tipo, PercRedBaseCalculo, Aliquota, PercMargemLucro, CategoriaFornecedor, CategoriaFornecedorItem, TipoCliForn, UsaPauta, ValorPauta, PercRedBaseCalculoSubst, TipoSimples, ICMSPercFCP, TipoAplicacao, ICMSSTBaseDupla, ICMSSTBaseDuplaIni, cBenef, ICMSMotivo FROM ICMSExcecoes WHERE EstadoDestino = ? ORDER BY Prioridade, GrupoOrigemMercadoria DESC", _
        tICMSExcecao.iGrupoOrigemMercadoria, tICMSExcecao.sEstadoOrigem, tICMSExcecao.sCategoriaProduto, tICMSExcecao.sCategoriaProdutoItem, tICMSExcecao.sCategoriaCliente, tICMSExcecao.sCategoriaClienteItem, tICMSExcecao.iTipo, tICMSExcecao.dPercRedBaseCalculo, tICMSExcecao.dAliquota, tICMSExcecao.dPercMargemLucro, tICMSExcecao.sCategoriaFornecedor, tICMSExcecao.sCategoriaFornecedorItem, tICMSExcecao.iTipoCliForn, tICMSExcecao.iUsaPauta, tICMSExcecao.dValorPauta, tICMSExcecao.dPercRedBaseCalculoSubst, tICMSExcecao.iTipoSimples, tICMSExcecao.dICMSPercFCP, tICMSExcecao.iTipoAplicacao, tICMSExcecao.iICMSSTBaseDupla, tICMSExcecao.dtICMSSTBaseDuplaIni, tICMSExcecao.scBenef, tICMSExcecao.iICMSMotivo, sEstado)
    If lErro <> AD_SQL_SUCESSO Then Error 27273
    '###########################################################
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27274
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objICMSExcecao = New ClassICMSExcecao
        
        objICMSExcecao.iGrupoOrigemMercadoria = tICMSExcecao.iGrupoOrigemMercadoria
        objICMSExcecao.sEstadoOrigem = tICMSExcecao.sEstadoOrigem
        objICMSExcecao.sEstadoDestino = sEstado
        objICMSExcecao.dAliquota = tICMSExcecao.dAliquota
        objICMSExcecao.dPercMargemLucro = tICMSExcecao.dPercMargemLucro
        objICMSExcecao.dPercRedBaseCalculo = tICMSExcecao.dPercRedBaseCalculo
        objICMSExcecao.iTipo = tICMSExcecao.iTipo
        objICMSExcecao.sCategoriaCliente = tICMSExcecao.sCategoriaCliente
        objICMSExcecao.sCategoriaClienteItem = tICMSExcecao.sCategoriaClienteItem
        objICMSExcecao.sCategoriaProduto = tICMSExcecao.sCategoriaProduto
        objICMSExcecao.sCategoriaProdutoItem = tICMSExcecao.sCategoriaProdutoItem
        
        '###########################################################
        'Inserido por Wagner 29/09/05
        objICMSExcecao.sCategoriaFornecedor = tICMSExcecao.sCategoriaFornecedor
        objICMSExcecao.sCategoriaFornecedorItem = tICMSExcecao.sCategoriaFornecedorItem
        objICMSExcecao.dPercRedBaseCalculoSubst = tICMSExcecao.dPercRedBaseCalculoSubst
        '###########################################################
        objICMSExcecao.iTipoCliForn = tICMSExcecao.iTipoCliForn
        objICMSExcecao.iUsaPauta = tICMSExcecao.iUsaPauta
        objICMSExcecao.dValorPauta = tICMSExcecao.dValorPauta
        objICMSExcecao.iTipoSimples = tICMSExcecao.iTipoSimples
        
        objICMSExcecao.dICMSPercFCP = tICMSExcecao.dICMSPercFCP
        objICMSExcecao.iTipoAplicacao = tICMSExcecao.iTipoAplicacao

        objICMSExcecao.iICMSSTBaseDupla = tICMSExcecao.iICMSSTBaseDupla
        objICMSExcecao.dtICMSSTBaseDuplaIni = tICMSExcecao.dtICMSSTBaseDuplaIni
        objICMSExcecao.scBenef = tICMSExcecao.scBenef
        objICMSExcecao.iICMSMotivo = tICMSExcecao.iICMSMotivo

        colExcecoes.Add objICMSExcecao
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27275
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    ICMSExcecoes_Le_Estado = SUCESSO

    Exit Function

Erro_ICMSExcecoes_Le_Estado:

    ICMSExcecoes_Le_Estado = Err

    Select Case Err

        Case 27272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27273, 27274, 27275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_EXCECOES", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153658)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TiposTribICMS_Le_Todos(ByVal colTiposTribICMS As Collection) As Long
'preenche colTiposTribICMS com objetos ClassTipoTribICMS a partir da tabela TiposTribICMS
Dim lErro As Long, lComando As Long
Dim objTipoTribICMS As ClassTipoTribICMS
Dim iPermiteAliquota As Integer, iPermiteMargLucro As Integer, iPermiteReducaoBase As Integer, iTipo As Integer, sDescricao As String
Dim iPermiteST As Integer, iPermitePercBaseOperProp As Integer, iPermiteUFICMSSTDevido As Integer
Dim iPermiteVlrICMSIsento As Integer, iPermiteMotDesoneracao As Integer, iPermiteSTRetUFRem As Integer
Dim iPermiteSTUFDest As Integer, iPermiteSTRetAnt As Integer, iTipoTribCST As Integer, iVersaoNFE As Integer

On Error GoTo Erro_TiposTribICMS_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27276
    
    sDescricao = String(STRING_TIPO_ICMS_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, PermiteAliquota, PermiteMargLucro, PermiteReducaoBase, PermiteST, PermitePercBaseOperProp, PermiteUFICMSSTDevido, PermiteVlrICMSIsento, PermiteMotDesoneracao,PermiteSTRetUFRem , PermiteSTUFDest, PermiteSTRetAnt, TipoTribCST, VersaoNFE FROM TiposTribICMS ORDER BY TipoTribCST", _
        iTipo, sDescricao, iPermiteAliquota, iPermiteMargLucro, iPermiteReducaoBase, iPermiteST, iPermitePercBaseOperProp, iPermiteUFICMSSTDevido, iPermiteVlrICMSIsento, iPermiteMotDesoneracao, iPermiteSTRetUFRem, iPermiteSTUFDest, iPermiteSTRetAnt, iTipoTribCST, iVersaoNFE)
    If lErro <> AD_SQL_SUCESSO Then Error 27277
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27278
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objTipoTribICMS = New ClassTipoTribICMS

        objTipoTribICMS.iPermiteAliquota = iPermiteAliquota
        objTipoTribICMS.iPermiteMargLucro = iPermiteMargLucro
        objTipoTribICMS.iPermiteReducaoBase = iPermiteReducaoBase
        objTipoTribICMS.iTipo = iTipo
        objTipoTribICMS.sDescricao = sDescricao
        objTipoTribICMS.iPermiteST = iPermiteST
        objTipoTribICMS.iPermitePercBaseOperProp = iPermitePercBaseOperProp
        objTipoTribICMS.iPermiteUFICMSSTDevido = iPermiteUFICMSSTDevido
        objTipoTribICMS.iPermiteVlrICMSIsento = iPermiteVlrICMSIsento
        objTipoTribICMS.iPermiteMotDesoneracao = iPermiteMotDesoneracao
        objTipoTribICMS.iPermiteSTRetUFRem = iPermiteSTRetUFRem
        objTipoTribICMS.iPermiteSTUFDest = iPermiteSTUFDest
        objTipoTribICMS.iPermiteSTRetAnt = iPermiteSTRetAnt
        objTipoTribICMS.iTipoTribCST = iTipoTribCST
        objTipoTribICMS.iVersaoNFE = iVersaoNFE
        
        colTiposTribICMS.Add objTipoTribICMS
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27279
    
    Loop
       
    lErro = Comando_Fechar(lComando)
    
    TiposTribICMS_Le_Todos = SUCESSO

    Exit Function

Erro_TiposTribICMS_Le_Todos:

    TiposTribICMS_Le_Todos = Err

    Select Case Err

        Case 27276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27277, 27278, 27279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBICMS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153659)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoTribICMS_Le(ByVal objTipoTribICMS As ClassTipoTribICMS) As Long
'Lê o item da tabela TipoTributacaoICMS com o Tipo passado em objTipoTributacaoICMS

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim iPermiteAliquota As Integer
Dim iPermiteMargLucro As Integer
Dim iPermiteReducaoBase As Integer
Dim iColunaNoLivroEntrada As Integer
Dim iColunaNoLivroSaida As Integer
Dim iTipoTribCST As Integer
Dim iPermiteST As Integer, iPermitePercBaseOperProp As Integer, iPermiteUFICMSSTDevido As Integer
Dim iPermiteVlrICMSIsento As Integer, iPermiteMotDesoneracao As Integer, iPermiteSTRetUFRem As Integer
Dim iPermiteSTUFDest As Integer, iPermiteSTRetAnt As Integer, iVersaoNFE As Integer

On Error GoTo Erro_TipoTribICMS_Le

    sDescricao = String(STRING_TIPO_ICMS_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21531

    'Lê a tabela de TipoTributacaoICMS
    lErro = Comando_Executar(lComando, "SELECT Descricao, PermiteAliquota, PermiteMargLucro, PermiteReducaoBase, ColunaNoLivroEntrada, ColunaNoLivroSaida, TipoTribCST, PermiteST, PermitePercBaseOperProp, PermiteUFICMSSTDevido, PermiteVlrICMSIsento, PermiteMotDesoneracao,PermiteSTRetUFRem , PermiteSTUFDest, PermiteSTRetAnt, VersaoNFE FROM TiposTribICMS WHERE Tipo = ?", _
    sDescricao, iPermiteAliquota, iPermiteMargLucro, iPermiteReducaoBase, iColunaNoLivroEntrada, iColunaNoLivroSaida, iTipoTribCST, iPermiteST, iPermitePercBaseOperProp, iPermiteUFICMSSTDevido, iPermiteVlrICMSIsento, iPermiteMotDesoneracao, iPermiteSTRetUFRem, iPermiteSTUFDest, iPermiteSTRetAnt, iVersaoNFE, objTipoTribICMS.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 21532

    'Lê o primeiro item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21533

    If lErro = AD_SQL_SEM_DADOS Then Error 21534

    'Transfere os dados para objTipoTribICMS
    objTipoTribICMS.sDescricao = sDescricao
    objTipoTribICMS.iPermiteAliquota = iPermiteAliquota
    objTipoTribICMS.iPermiteReducaoBase = iPermiteReducaoBase
    objTipoTribICMS.iPermiteMargLucro = iPermiteMargLucro
    objTipoTribICMS.iColunaNoLivroEntrada = iColunaNoLivroEntrada
    objTipoTribICMS.iColunaNoLivroSaida = iColunaNoLivroSaida
    objTipoTribICMS.iTipoTribCST = iTipoTribCST
    objTipoTribICMS.iPermiteST = iPermiteST
    objTipoTribICMS.iPermitePercBaseOperProp = iPermitePercBaseOperProp
    objTipoTribICMS.iPermiteUFICMSSTDevido = iPermiteUFICMSSTDevido
    objTipoTribICMS.iPermiteVlrICMSIsento = iPermiteVlrICMSIsento
    objTipoTribICMS.iPermiteMotDesoneracao = iPermiteMotDesoneracao
    objTipoTribICMS.iPermiteSTRetUFRem = iPermiteSTRetUFRem
    objTipoTribICMS.iPermiteSTUFDest = iPermiteSTUFDest
    objTipoTribICMS.iPermiteSTRetAnt = iPermiteSTRetAnt
    objTipoTribICMS.iVersaoNFE = iVersaoNFE
    
    Call Comando_Fechar(lComando)

    TipoTribICMS_Le = SUCESSO

    Exit Function

Erro_TipoTribICMS_Le:

    TipoTribICMS_Le = Err

    Select Case Err

        Case 21531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21532, 21533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBICMS", Err)

        Case 21534

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153660)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ICMSExcecao_Le(ByVal objICMSExcecoes As ClassICMSExcecao) As Long
'Lê um item da tabela ICMSExcecoes, a partir dos dados passados em objICMSExcecoes

Dim lErro As Long
Dim lComando As Long
Dim iTipo As Integer
Dim dAliquota As Double
Dim iPrioridade As Integer
Dim sFundamentacao As String
Dim dPercMargemLucro As Double
Dim dPercRedBaseCalculo As Double, iUsaPauta As Integer, dValorPauta As Double
Dim dPercRedBaseCalculoSubst As Double
Dim iTipoCliForn As Integer 'Inserido por Wagner
Dim iTipoSimples As Integer
Dim dICMSPercFCP As Double
Dim iICMSSTBaseDupla As Integer
Dim dtICMSSTBaseDuplaIni As Date
Dim scBenef As String
Dim iICMSMotivo As Integer

On Error GoTo Erro_ICMSExcecao_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21537

    sFundamentacao = String(STRING_EXCECAO_TRIB_FUNDAMENTACAO, 0)
    scBenef = String(20, 0)
    '#############################################
    'Alterado por Wagner 29/09/05
    'Lê a tabela
    lErro = Comando_Executar(lComando, "SELECT Aliquota, PercMargemLucro, PercRedBaseCalculo, Fundamentacao, Tipo, Prioridade, TipoCliForn, UsaPauta, ValorPauta, PercRedBaseCalculoSubst, TipoSimples, ICMSPercFCP, ICMSSTBaseDupla, ICMSSTBaseDuplaIni, cBenef, ICMSMotivo FROM ICMSExcecoes WHERE EstadoDestino = ? AND CategoriaProduto = ? AND CategoriaProdutoItem = ? AND CategoriaCliente = ? AND CategoriaClienteItem = ? AND EstadoOrigem = ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem = ? AND GrupoOrigemMercadoria = ? AND TipoAplicacao = ?", _
        dAliquota, dPercMargemLucro, dPercRedBaseCalculo, sFundamentacao, iTipo, iPrioridade, iTipoCliForn, iUsaPauta, dValorPauta, dPercRedBaseCalculoSubst, iTipoSimples, dICMSPercFCP, iICMSSTBaseDupla, dtICMSSTBaseDuplaIni, scBenef, iICMSMotivo, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.iTipoAplicacao)
    If lErro <> AD_SQL_SUCESSO Then Error 21538
    '#############################################

    'Lê o primeiro item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21541

    If lErro = AD_SQL_SEM_DADOS Then Error 21542

    'Transfere dados para objICMSExcecoes
    objICMSExcecoes.sFundamentacao = sFundamentacao
    objICMSExcecoes.iTipo = iTipo
    objICMSExcecoes.dAliquota = dAliquota
    objICMSExcecoes.dPercMargemLucro = dPercMargemLucro
    objICMSExcecoes.dPercRedBaseCalculo = dPercRedBaseCalculo
    objICMSExcecoes.iPrioridade = iPrioridade
    objICMSExcecoes.iTipoCliForn = iTipoCliForn 'Inserido por Wagner
    objICMSExcecoes.iUsaPauta = iUsaPauta
    objICMSExcecoes.dValorPauta = dValorPauta
    objICMSExcecoes.dPercRedBaseCalculoSubst = dPercRedBaseCalculoSubst
    objICMSExcecoes.iTipoSimples = iTipoSimples
    objICMSExcecoes.dICMSPercFCP = dICMSPercFCP
    objICMSExcecoes.iICMSSTBaseDupla = iICMSSTBaseDupla
    objICMSExcecoes.dtICMSSTBaseDuplaIni = dtICMSSTBaseDuplaIni
    objICMSExcecoes.scBenef = scBenef
    objICMSExcecoes.iICMSMotivo = iICMSMotivo
    
    'Fecha comando
    Call Comando_Fechar(lComando)

    ICMSExcecao_Le = SUCESSO

    Exit Function

Erro_ICMSExcecao_Le:

    ICMSExcecao_Le = Err

    Select Case Err

        Case 21537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21538, 21541, 21707, 21708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_EXCECOES", Err)

        Case 21542

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153661)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoTribIPI_Le(ByVal objTipoTribIPI As ClassTipoTribIPI) As Long
'Lê o item da tabela TipoTributacaoIPI com o Tipo passado em objTipoTributacaoIPI

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim iPermiteAliquota As Integer
Dim iPermiteReducaoBase As Integer, iCSTEntrada As Integer, iCSTSaida As Integer, iTipoCalculo As Integer
Dim iColunaNoLivro As Integer, iVersaoNFE As Integer, iAtivo As Integer

On Error GoTo Erro_TipoTribIPI_Le

    sDescricao = String(STRING_TIPO_IPI_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21617

    'Lê a tabela de TipoTributacaoIPI
    lErro = Comando_Executar(lComando, "SELECT Descricao, PermiteAliquota, PermiteReducaoBase, ColunaNoLivro, VersaoNFe, Ativo, CSTEntrada, CSTSaida, TipoCalculo FROM TiposTribIPI WHERE Tipo = ?", sDescricao, iPermiteAliquota, iPermiteReducaoBase, iColunaNoLivro, iVersaoNFE, iAtivo, iCSTEntrada, iCSTSaida, iTipoCalculo, objTipoTribIPI.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 21618

    'Lê o primeiro item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21619

    If lErro = AD_SQL_SEM_DADOS Then Error 21620

    'Transfere os dados para objTipoTribIPI
    objTipoTribIPI.sDescricao = sDescricao
    objTipoTribIPI.iPermiteAliquota = iPermiteAliquota
    objTipoTribIPI.iPermiteReducaoBase = iPermiteReducaoBase
    objTipoTribIPI.iColunaNoLivro = iColunaNoLivro
    objTipoTribIPI.iVersaoNFE = iVersaoNFE
    objTipoTribIPI.iAtivo = iAtivo
    objTipoTribIPI.iCSTEntrada = iCSTEntrada
    objTipoTribIPI.iCSTSaida = iCSTSaida
    objTipoTribIPI.iTipoCalculo = iTipoCalculo

    Call Comando_Fechar(lComando)

    TipoTribIPI_Le = SUCESSO

    Exit Function

Erro_TipoTribIPI_Le:

    TipoTribIPI_Le = Err

    Select Case Err

        Case 21617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21618, 21619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBIPI", Err)

        Case 21620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153662)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposTribIPI_Le_Todos(ByVal colTiposTribIPI As Collection) As Long
'preenche colTiposTribIPI com objetos ClassTipoTribIPI a partir da tabela TiposTribIPI
Dim lErro As Long, lComando As Long
Dim objTipoTribIPI As ClassTipoTribIPI
Dim iPermiteAliquota As Integer, iPermiteReducaoBase As Integer, iTipo As Integer, sDescricao As String
Dim iVersaoNFE As Integer, iCSTEntrada As Integer, iCSTSaida As Integer, iTipoCalculo As Integer

On Error GoTo Erro_TiposTribIPI_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27511
    
    sDescricao = String(STRING_TIPO_IPI_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, PermiteAliquota, PermiteReducaoBase, VersaoNFe, CSTEntrada, CSTSaida, TipoCalculo FROM TiposTribIPI WHERE Ativo = 1 ORDER BY CSTSaida", _
        iTipo, sDescricao, iPermiteAliquota, iPermiteReducaoBase, iVersaoNFE, iCSTEntrada, iCSTSaida, iTipoCalculo)
    If lErro <> AD_SQL_SUCESSO Then Error 27512
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27513
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objTipoTribIPI = New ClassTipoTribIPI

        objTipoTribIPI.iPermiteAliquota = iPermiteAliquota
        objTipoTribIPI.iPermiteReducaoBase = iPermiteReducaoBase
        objTipoTribIPI.iTipo = iTipo
        objTipoTribIPI.sDescricao = sDescricao
        objTipoTribIPI.iVersaoNFE = iVersaoNFE
        objTipoTribIPI.iCSTEntrada = iCSTEntrada
        objTipoTribIPI.iCSTSaida = iCSTSaida
        objTipoTribIPI.iTipoCalculo = iTipoCalculo
        
        colTiposTribIPI.Add objTipoTribIPI
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27514
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    TiposTribIPI_Le_Todos = SUCESSO

    Exit Function

Erro_TiposTribIPI_Le_Todos:

    TiposTribIPI_Le_Todos = Err

    Select Case Err

        Case 27511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27512, 27513, 27514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBIPI", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153663)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function IPIExcecoes_Le_Todas(ByVal colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassIPIExcecoes
Dim lErro As Long, lComando As Long
Dim objIPIExcecao As ClassIPIExcecao, tIPIExcecao As typeIPIExcecao
On Error GoTo Erro_IPIExcecoes_Le_Todas

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27526
    
    tIPIExcecao.sCategoriaCliente = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    tIPIExcecao.sCategoriaClienteItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    tIPIExcecao.sCategoriaProduto = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tIPIExcecao.sCategoriaProdutoItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT CategoriaProduto, CategoriaProdutoItem, CategoriaCliente, CategoriaClienteItem, Tipo, PercRedBaseCalculo, Aliquota, TipoCalculo, AliquotaRS FROM IPIExcecoes", _
        tIPIExcecao.sCategoriaProduto, tIPIExcecao.sCategoriaProdutoItem, tIPIExcecao.sCategoriaCliente, tIPIExcecao.sCategoriaClienteItem, tIPIExcecao.iTipo, tIPIExcecao.dPercRedBaseCalculo, tIPIExcecao.dAliquota, tIPIExcecao.iTipoCalculo, tIPIExcecao.dAliquotaRS)
    If lErro <> AD_SQL_SUCESSO Then Error 27527
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27528
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objIPIExcecao = New ClassIPIExcecao
        
        objIPIExcecao.dAliquota = tIPIExcecao.dAliquota
        objIPIExcecao.dPercRedBaseCalculo = tIPIExcecao.dPercRedBaseCalculo
        objIPIExcecao.iTipo = tIPIExcecao.iTipo
        objIPIExcecao.sCategoriaCliente = tIPIExcecao.sCategoriaCliente
        objIPIExcecao.sCategoriaClienteItem = tIPIExcecao.sCategoriaClienteItem
        objIPIExcecao.sCategoriaProduto = tIPIExcecao.sCategoriaProduto
        objIPIExcecao.sCategoriaProdutoItem = tIPIExcecao.sCategoriaProdutoItem
        objIPIExcecao.iTipoCalculo = tIPIExcecao.iTipoCalculo
        objIPIExcecao.dAliquotaRS = tIPIExcecao.dAliquotaRS
        
        colExcecoes.Add objIPIExcecao
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27529
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    IPIExcecoes_Le_Todas = SUCESSO

    Exit Function

Erro_IPIExcecoes_Le_Todas:

    IPIExcecoes_Le_Todas = Err

    Select Case Err

        Case 27526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27527, 27528, 27529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPI_EXCECOES", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153664)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function IPIExcecoes_Le_FilialCliente(ByVal lCliente As Long, ByVal iFilial As Integer, ByVal colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassIPIExcecoes validos para a filial cliente, ordenandos pelo campo "prioridade"
Dim lErro As Long, lComando As Long
Dim objIPIExcecao As ClassIPIExcecao, tIPIExcecao As typeIPIExcecao
On Error GoTo Erro_IPIExcecoes_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27591
    
    tIPIExcecao.sCategoriaCliente = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    tIPIExcecao.sCategoriaClienteItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    tIPIExcecao.sCategoriaProduto = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tIPIExcecao.sCategoriaProdutoItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT CategoriaProduto, CategoriaProdutoItem, CategoriaCliente, CategoriaClienteItem, Tipo, PercRedBaseCalculo, Aliquota, TipoCalculo, AliquotaRS FROM IPIExcecoes WHERE CategoriaCliente = '' OR EXISTS (SELECT Filial FROM FilialClienteCategorias WHERE Cliente = ? AND Filial = ? AND Categoria = IPIExcecoes.CategoriaCliente AND Item = IPIExcecoes.CategoriaClienteItem) ORDER BY Prioridade", _
        tIPIExcecao.sCategoriaProduto, tIPIExcecao.sCategoriaProdutoItem, tIPIExcecao.sCategoriaCliente, tIPIExcecao.sCategoriaClienteItem, tIPIExcecao.iTipo, tIPIExcecao.dPercRedBaseCalculo, tIPIExcecao.dAliquota, tIPIExcecao.iTipoCalculo, tIPIExcecao.dAliquotaRS, lCliente, iFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 27592
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27593
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objIPIExcecao = New ClassIPIExcecao
        
        objIPIExcecao.dAliquota = tIPIExcecao.dAliquota
        objIPIExcecao.dPercRedBaseCalculo = tIPIExcecao.dPercRedBaseCalculo
        objIPIExcecao.iTipo = tIPIExcecao.iTipo
        objIPIExcecao.sCategoriaCliente = tIPIExcecao.sCategoriaCliente
        objIPIExcecao.sCategoriaClienteItem = tIPIExcecao.sCategoriaClienteItem
        objIPIExcecao.sCategoriaProduto = tIPIExcecao.sCategoriaProduto
        objIPIExcecao.sCategoriaProdutoItem = tIPIExcecao.sCategoriaProdutoItem
        objIPIExcecao.iTipoCalculo = tIPIExcecao.iTipoCalculo
        objIPIExcecao.dAliquotaRS = tIPIExcecao.dAliquotaRS
        
        colExcecoes.Add objIPIExcecao
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27594
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    IPIExcecoes_Le_FilialCliente = SUCESSO

    Exit Function

Erro_IPIExcecoes_Le_FilialCliente:

    IPIExcecoes_Le_FilialCliente = Err

    Select Case Err

        Case 27591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27592, 27593, 27594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPI_EXCECOES", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153665)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoTributacao_Le(ByVal objTributacaoTipo As ClassTipoDeTributacaoMovto) As Long
'Preenche objTributacaoTipo a partir do bd p/registro identificado por objTributacaoTipo.iTipo
'Retorna 27259 se nao encontrar
Dim lErro As Long, lComando As Long
Dim tTributacaoTipo As typeTributacaoTipo

On Error GoTo Erro_TipoTributacao_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27256
    
    tTributacaoTipo.sDescricao = String(TIPO_TRIBUTACAO_DESCRICAO, 0)
    tTributacaoTipo.sNatBCCred = String(STRING_NATBCCRED, 0)
    tTributacaoTipo.sIPICodEnq = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, Entrada, ICMSIncide, ICMSTipo, ICMSBaseComIPI, ICMSCredita, IPIIncide, IPITipo, IPIFrete, IPIDestaca, IPICredita, ISSIncide, IRIncide, IRAliquota, INSSAliquota, INSSRetencaoMinima, INSSIncide, PISCredita, PISRetencao, ISSRetencao, COFINSCredita, COFINSRetencao, CSLLRetencao, ISSTipo , PISTipo, COFINSTipo, ICMSSimplesTipo, RegimeTributario, NatBCCred, ISSIndExigibilidade, IPICodEnq  FROM TiposDeTributacaoMovto WHERE Tipo = ?", _
        tTributacaoTipo.sDescricao, tTributacaoTipo.iEntrada, tTributacaoTipo.iICMSIncide, tTributacaoTipo.iICMSTipo, tTributacaoTipo.iICMSBaseComIPI, tTributacaoTipo.iICMSCredita, tTributacaoTipo.iIPIIncide, tTributacaoTipo.iIPITipo, tTributacaoTipo.iIPIFrete, tTributacaoTipo.iIPIDestaca, tTributacaoTipo.iIPICredita, tTributacaoTipo.iISSIncide, tTributacaoTipo.iIRIncide, tTributacaoTipo.dIRAliquota, tTributacaoTipo.dINSSAliquota, tTributacaoTipo.dINSSRetencaoMinima, tTributacaoTipo.iINSSIncide, tTributacaoTipo.iPISCredita, tTributacaoTipo.iPISRetencao, tTributacaoTipo.iISSRetencao, tTributacaoTipo.iCOFINSCredita, tTributacaoTipo.iCOFINSRetencao, tTributacaoTipo.iCSLLRetencao, tTributacaoTipo.iISSTipo, tTributacaoTipo.iPISTipo, tTributacaoTipo.iCOFINSTipo, tTributacaoTipo.iICMSSimplesTipo, tTributacaoTipo.iRegimeTributario, tTributacaoTipo.sNatBCCred, tTributacaoTipo.iISSIndExigibilidade, tTributacaoTipo.sIPICodEnq, objTributacaoTipo.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 27257
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27258
        
    If lErro <> AD_SQL_SUCESSO Then Error 27259
    
    With tTributacaoTipo
        objTributacaoTipo.iICMSBaseComIPI = .iICMSBaseComIPI
        objTributacaoTipo.iICMSCredita = .iICMSCredita
        objTributacaoTipo.iICMSIncide = .iICMSIncide
        objTributacaoTipo.iICMSTipo = .iICMSTipo
        objTributacaoTipo.iIPIDestaca = .iIPIDestaca
        objTributacaoTipo.iIPICredita = .iIPICredita
        objTributacaoTipo.iIPIFrete = .iIPIFrete
        objTributacaoTipo.iIPIIncide = .iIPIIncide
        objTributacaoTipo.iIPITipo = .iIPITipo
        objTributacaoTipo.iIRIncide = .iIRIncide
        objTributacaoTipo.iISSIncide = .iISSIncide
        objTributacaoTipo.iEntrada = .iEntrada
        objTributacaoTipo.sDescricao = .sDescricao
        objTributacaoTipo.dIRAliquota = .dIRAliquota
        objTributacaoTipo.dINSSAliquota = .dINSSAliquota
        objTributacaoTipo.dINSSRetencaoMinima = .dINSSRetencaoMinima
        objTributacaoTipo.iINSSIncide = .iINSSIncide
        objTributacaoTipo.iPISCredita = .iPISCredita
        objTributacaoTipo.iPISRetencao = .iPISRetencao
        objTributacaoTipo.iISSRetencao = .iISSRetencao
        objTributacaoTipo.iCOFINSCredita = .iCOFINSCredita
        objTributacaoTipo.iCOFINSRetencao = .iCOFINSRetencao
        objTributacaoTipo.iCSLLRetencao = .iCSLLRetencao
        objTributacaoTipo.iISSTipo = .iISSTipo
        objTributacaoTipo.iPISTipo = .iPISTipo
        objTributacaoTipo.iCOFINSTipo = .iCOFINSTipo
        objTributacaoTipo.iICMSSimplesTipo = .iICMSSimplesTipo
        objTributacaoTipo.iRegimeTributario = .iRegimeTributario
        objTributacaoTipo.sNatBCCred = .sNatBCCred
        objTributacaoTipo.iISSIndExigibilidade = .iISSIndExigibilidade
        objTributacaoTipo.sIPICodEnq = .sIPICodEnq
    End With
    
    lErro = Comando_Fechar(lComando)
    
    TipoTributacao_Le = SUCESSO

    Exit Function

Erro_TipoTributacao_Le:

    TipoTributacao_Le = Err

    Select Case Err

        Case 27256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27257, 27258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", Err, objTributacaoTipo.iTipo)
        
        Case 27259 'tipo nao encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153666)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PadraoTribSaida_Le(ByVal objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Lê o Padrão de Tributação Saída passado como parâmetro em objPadraoTribSaida

Dim lErro As Long
Dim lComando As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribSaida_Le

    'Inicializa o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33382

    'Pesquisa no BD o Padrão de Tributação
    lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaFilialCliente = ? AND ItemCategoria = ?", iTipoTributacaoPadrao, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33383

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33384

    'Não encontrou o Padrão de Tributação
    If lErro = AD_SQL_SEM_DADOS Then Error 33385

    'Transfere o dado para objPadraoTribSaida
    objPadraoTribSaida.iTipoTributacaoPadrao = iTipoTributacaoPadrao

    'Libera o comando
    Call Comando_Fechar(lComando)

    PadraoTribSaida_Le = SUCESSO

    Exit Function

Erro_PadraoTribSaida_Le:

    PadraoTribSaida_Le = Err

    Select Case Err

        Case 33382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33383, 33384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBUTACAO1", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33385
            'Não encontrou o Padrão de Tributação. Erro tratado na rotina chamada.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153667)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PadraoTribEntrada_Le(ByVal objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Lê o Padrão de Tributação Entrada passado como parâmetro em objPadraoTribEnt

Dim lErro As Long
Dim lComando As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribEntrada_Le

    'Inicializa o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33480

    'Pesquisa no BD o Padrão de Tributação de Entrada
    lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaProduto = ? AND ItemCategoria = ?", iTipoTributacaoPadrao, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33481

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33482

    'Não encontrou o Padrão de Tributação Entrada
    If lErro = AD_SQL_SEM_DADOS Then Error 33483

    'Transfere o dado para objPadraoTribSaida
    objPadraoTribEnt.iTipoTributacaoPadrao = iTipoTributacaoPadrao

    'Libera o comando
    Call Comando_Fechar(lComando)

    PadraoTribEntrada_Le = SUCESSO

    Exit Function

Erro_PadraoTribEntrada_Le:

    PadraoTribEntrada_Le = Err

    Select Case Err

        Case 33480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33481, 33482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33483
            'Não encontrou o Padrão de Tributação Entrada. Erro tratado na rotina chamada.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153668)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ICMSAliqExternas_Le_EstadoDestino(ByVal sSiglaDestino As String, ByVal colAliqExternas As Collection) As Long
'carrega as aliquotas externas que chegam a um estado (sSiglaDestino)
Dim lErro As Long, lComando As Long, sSiglaOrigem As String, dAliquota As Double
Dim objICMSAliqExterna As ClassICMSAliqExterna
On Error GoTo Erro_ICMSAliqExternas_Le_EstadoDestino

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27863
    
    sSiglaOrigem = String(STRING_ESTADO, 0)
    
    'ler ICMSAliquotasExternas
    lErro = Comando_Executar(lComando, "SELECT SiglaOrig, Aliquota FROM ICMSAliquotasExternas WHERE SiglaDest = ?", sSiglaOrigem, dAliquota, sSiglaDestino)
    If lErro <> AD_SQL_SUCESSO Then Error 27864
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27865
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objICMSAliqExterna = New ClassICMSAliqExterna
        
        objICMSAliqExterna.dAliquota = dAliquota
        objICMSAliqExterna.sSiglaDest = sSiglaDestino
        objICMSAliqExterna.sSiglaOrig = sSiglaOrigem
        
        colAliqExternas.Add objICMSAliqExterna, sSiglaOrigem
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27866
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    ICMSAliqExternas_Le_EstadoDestino = SUCESSO

    Exit Function

Erro_ICMSAliqExternas_Le_EstadoDestino:

    ICMSAliqExternas_Le_EstadoDestino = Err

    Select Case Err

        Case 27863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27864, 27865, 27866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_ALIQ_EXT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153669)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function IPIExcecao_Le(ByVal objIPIExcecoes As ClassIPIExcecao) As Long
'Lê um item da tabela IPIExcecoes, a partir dos dados passados em objIPIExcecoes

Dim lErro As Long
Dim lComando As Long
Dim iTipo As Integer, iTipoCalculo As Integer
Dim dAliquota As Double, dAliquotaRS As Double
Dim sFundamentacao As String
Dim dPercMargemLucro As Double
Dim dPercRedBaseCalculo As Double

On Error GoTo Erro_IPIExcecao_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21621

    sFundamentacao = String(150, 0)

    'Lê a tabela
    lErro = Comando_Executar(lComando, "SELECT Aliquota, PercMargemLucro, PercRedBaseCalculo, Fundamentacao, Tipo, TipoCalculo, AliquotaRS FROM IPIExcecoes WHERE CategoriaProduto = ? AND CategoriaProdutoItem = ? AND CategoriaCliente = ? AND CategoriaClienteItem = ? ", dAliquota, dPercMargemLucro, dPercRedBaseCalculo, sFundamentacao, iTipo, iTipoCalculo, dAliquotaRS, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)
    If lErro <> AD_SQL_SUCESSO Then Error 21622

    'Lê o primeiro item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21623

    If lErro = AD_SQL_SEM_DADOS Then Error 21624

    'Transfere dados para objIPIExcecoes
    objIPIExcecoes.iTipo = iTipo
    objIPIExcecoes.dAliquota = dAliquota
    objIPIExcecoes.sFundamentacao = sFundamentacao
    objIPIExcecoes.dPercMargemLucro = dPercMargemLucro
    objIPIExcecoes.dPercRedBaseCalculo = dPercRedBaseCalculo
    objIPIExcecoes.iTipoCalculo = iTipoCalculo
    objIPIExcecoes.dAliquotaRS = dAliquotaRS

    'Fecha comando
    Call Comando_Fechar(lComando)

    IPIExcecao_Le = SUCESSO

    Exit Function

Erro_IPIExcecao_Le:

    IPIExcecao_Le = Err

    Select Case Err

        Case 21621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21622, 21623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPIEXCECOES", Err)

        Case 21624

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153670)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ObterTribPadraoCli(iTipoTrib As Integer, ByVal sTipoDocSigla As String, ByVal sNatOp As String, ByVal lCliente As Long, ByVal iFilialCliente As Integer) As Long
'retorna em iTipoTrib o tipo de tributacao default p/o tipo do docto, natureza de operacao e filial cliente informados
'a pesquisa é feita na tabela PadroesTributacao, pegando um padrao correspondente a uma categoria associada a filialCliente
'Se nao achar, pode retornar um padrao generico, que independa do cliente. Se ainda assim nao achar nenhum registro retorna zero.

Dim lErro As Long, lComando As Long, objPadraoTribSaida As ClassPadraoTribSaida

On Error GoTo Erro_ObterTribPadraoCli
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27252
        
    If giSQLTipoOrdParam = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE (NaturezaOperacao = ? OR NaturezaOperacao = ?) AND (SiglaMovto = ? OR SiglaMovto = ?) AND ((CategoriaFilialCliente = ? AND ItemCategoria = ? ) OR (EXISTS (SELECT FilialClienteCategorias.Cliente FROM FilialClienteCategorias WHERE FilialClienteCategorias.Cliente = ? AND FilialClienteCategorias.Filial = ? AND FilialClienteCategorias.Categoria = PadroesTribSaida.CategoriaFilialCliente AND FilialClienteCategorias.Item = PadroesTribSaida.ItemCategoria))) ORDER BY CategoriaFilialCliente DESC, NaturezaOperacao DESC, SiglaMovto DESC", _
            iTipoTrib, sNatOp, "", sTipoDocSigla, "", "", "", lCliente, iFilialCliente)
    Else
        lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE (NaturezaOperacao = ? OR NaturezaOperacao = ?) AND (SiglaMovto = ? OR SiglaMovto = ?) AND ((CategoriaFilialCliente = ? AND ItemCategoria = ? ) OR (EXISTS (SELECT FilialClienteCategorias.Cliente FROM FilialClienteCategorias WHERE FilialClienteCategorias.Cliente = ? AND FilialClienteCategorias.Filial = ? AND FilialClienteCategorias.Categoria = PadroesTribSaida.CategoriaFilialCliente AND FilialClienteCategorias.Item = PadroesTribSaida.ItemCategoria))) ORDER BY CategoriaFilialCliente DESC, NaturezaOperacao DESC, SiglaMovto DESC", _
            iTipoTrib, lCliente, iFilialCliente, sNatOp, "", sTipoDocSigla, "", "", "")
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 27253
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27254
        
    'se nao achei nenhum padrao
    If lErro <> AD_SQL_SUCESSO Then iTipoTrib = 0
    
    lErro = Comando_Fechar(lComando)
    
    ObterTribPadraoCli = SUCESSO

    Exit Function

Erro_ObterTribPadraoCli:

    ObterTribPadraoCli = Err

    Select Case Err

        Case 27252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27253, 27254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROES_TRIBUTACAO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153671)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ObterTribPadraoProd(iTipoTrib As Integer, ByVal sTipoDocSigla As String, ByVal sNatOp As String, ByVal sProduto As String) As Long
'retorna em iTipoTrib o tipo de tributacao default p/o tipo do docto, natureza de operacao e produtos informados
'a pesquisa é feita na tabela PadroesTribEntrada, pegando um padrao correspondente a uma categoria associada ao produto
'Se nao achar, pode retornar um padrao generico, que independa do produto. Se ainda assim nao achar nenhum registro retorna zero.

Dim lErro As Long, lComando As Long, objPadraoTribSaida As ClassPadraoTribSaida

On Error GoTo Erro_ObterTribPadraoProd
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27887
    
    If giSQLTipoOrdParam = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE (NaturezaOperacao = ? OR NaturezaOperacao = ?) AND (SiglaMovto = ? OR SiglaMovto = ?) AND ((CategoriaProduto = ? AND ItemCategoria = ? ) OR (EXISTS (SELECT ProdutoCategoria.Produto FROM ProdutoCategoria WHERE ProdutoCategoria.Produto = ? AND ProdutoCategoria.Categoria = PadroesTribEntrada.CategoriaProduto AND ProdutoCategoria.Item = PadroesTribEntrada.ItemCategoria))) ORDER BY CategoriaProduto DESC, NaturezaOperacao DESC, SiglaMovto DESC", _
            iTipoTrib, sNatOp, "", sTipoDocSigla, "", "", "", sProduto)
    Else
        lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE (NaturezaOperacao = ? OR NaturezaOperacao = ?) AND (SiglaMovto = ? OR SiglaMovto = ?) AND ((CategoriaProduto = ? AND ItemCategoria = ? ) OR (EXISTS (SELECT ProdutoCategoria.Produto FROM ProdutoCategoria WHERE ProdutoCategoria.Produto = ? AND ProdutoCategoria.Categoria = PadroesTribEntrada.CategoriaProduto AND ProdutoCategoria.Item = PadroesTribEntrada.ItemCategoria))) ORDER BY CategoriaProduto DESC, NaturezaOperacao DESC, SiglaMovto DESC", _
            iTipoTrib, sProduto, sNatOp, "", sTipoDocSigla, "", "", "")
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 27888
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27889
        
    If lErro <> AD_SQL_SUCESSO Then iTipoTrib = 0
    
    lErro = Comando_Fechar(lComando)
    
    ObterTribPadraoProd = SUCESSO

    Exit Function

Erro_ObterTribPadraoProd:

    ObterTribPadraoProd = Err

    Select Case Err

        Case 27887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27888, 27889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROES_TRIBUTACAO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153672)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function NatOpPadrao_Le_Padrao(ByVal objNatOpPadrao As ClassNatOpPadrao) As Long
'obtem dados default p/tipo de operacao e tipo de atividade da empresa (comercio, industria,...)

Dim lErro As Long, lComando As Long, tNatOpPadrao As typeNatOpPadrao

On Error GoTo Erro_NatOpPadrao_Le_Padrao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56941
        
    tNatOpPadrao.sCFOPEmp = String(STRING_NATUREZAOP_CODIGO, 0)
    tNatOpPadrao.sCFOPExt = String(STRING_NATUREZAOP_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, TipoTribEmp, CFOPEmp, TipoAtividadeExt, TipoTribExt, CFOPExt FROM NatOpPadrao WHERE TipoOperacao = ? AND TipoAtividadeEmp = ? AND Padrao = ?", _
        tNatOpPadrao.lCodigo, tNatOpPadrao.iTipoTribEmp, tNatOpPadrao.sCFOPEmp, tNatOpPadrao.iTipoAtividadeExt, tNatOpPadrao.iTipoTribExt, tNatOpPadrao.sCFOPExt, objNatOpPadrao.iTipoOperacao, objNatOpPadrao.iTipoAtividadeEmp, 1)
    If lErro <> AD_SQL_SUCESSO Then Error 56942
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56943
        
    'se nao achei nenhum padrao
    If lErro <> AD_SQL_SUCESSO Then Error 56944
    
    With objNatOpPadrao
    
        .lCodigo = tNatOpPadrao.lCodigo
        .iTipoTribEmp = tNatOpPadrao.iTipoTribEmp
        .sCFOPEmp = tNatOpPadrao.sCFOPEmp
        .iTipoAtividadeExt = tNatOpPadrao.iTipoAtividadeExt
        .iTipoTribExt = tNatOpPadrao.iTipoTribExt
        .sCFOPExt = tNatOpPadrao.sCFOPExt
        .iPadrao = 1
        
    End With
    
    lErro = Comando_Fechar(lComando)
    
    NatOpPadrao_Le_Padrao = SUCESSO
     
    Exit Function
    
Erro_NatOpPadrao_Le_Padrao:

    NatOpPadrao_Le_Padrao = Err
     
    Select Case Err
          
        Case 56941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 56942, 56943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATOPPADRAO", Err)
        
        Case 56944 'nao achou
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153673)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

' *** 11/04/02 - INÍCIO Luiz G.F.Nogueira ***
Public Function ICMSAliqExternas_Le_Origem_Destino(objICMSAliqExterna As ClassICMSAliqExterna) As Long
'Lê a alíquota externa de ICMS para o par Origem, Destino passados como parâmetros

Dim lComando As Long
Dim lErro As Long
Dim dAliquota As Double
Dim sNomeTabela As String

On Error GoTo Erro_ICMSAliqExternas_Le_Origem_Destino
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94991
    
    'Guarda o nome da tabela que será lida
    sNomeTabela = "ICMSAliquotasExternas"
    
    'Procura no BD um registro que tenha a SiglaOrig e a SiglaDest passadas como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Aliquota FROM ICMSAliquotasExternas WHERE SiglaOrig=? AND SiglaDest=?", dAliquota, objICMSAliqExterna.sSiglaOrig, objICMSAliqExterna.sSiglaDest)
    If lErro <> AD_SQL_SUCESSO Then gError 94992
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94993
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 94994
    
    'Guarda no obj os dados lidos
    objICMSAliqExterna.dAliquota = dAliquota
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ICMSAliqExternas_Le_Origem_Destino = SUCESSO
    
    Exit Function
    
Erro_ICMSAliqExternas_Le_Origem_Destino:

    ICMSAliqExternas_Le_Origem_Destino = gErr
    
    Select Case gErr
    
        Case 94991
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94992, 94993
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GENERICO", gErr, sNomeTabela)
            
        Case 94994 'SEM DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153674)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
' *** 11/04/02 - FIM Luiz G.F.Nogueira ***
Public Function Define_SituacaoTributaria_Produto(sSituacaoTrib As String, ByVal sProduto As String, ByVal iTipoTribICMS As Integer, ByVal iTipoTribIPI As Integer) As Long
'Le a Situacao Tributaria (ICMS E IPI) - Retornando o tipo desejado (Tabela - Ecolub).
'O produto foi passado a pedido de Mario para facilitar posteriores mudancas.

Dim lErro As Long
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTribICMS As New ClassTipoTribICMS

On Error GoTo Erro_Define_SituacaoTributaria_Produto

    'Le o tipo de tributacao com o Tipo de icms passado
    objTipoTribICMS.iTipo = iTipoTribICMS
    lErro = CF("TipoTribICMS_Le", objTipoTribICMS)
    If lErro <> SUCESSO And lErro <> 21534 Then gError 108760
    
    'Se nao encontrou => Erro
    If lErro = 21534 Then gError 108761
    
    'Le o tipo de tributacao com o Tipo de ipi passado
    objTipoTribIPI.iTipo = iTipoTribIPI
    lErro = CF("TipoTribIPI_Le", objTipoTribIPI)
    If lErro <> SUCESSO And lErro <> 21620 Then gError 108762
    
    'Se nao encontrou => Erro
    If lErro = 21620 Then gError 108763
    
    'Apenas por seguranca ...
    sSituacaoTrib = ""
    
    'Definicao do Tipo A - Tributado de ICMS/IPI
    If objTipoTribICMS.iTipo <> ICMS_TIPO_NAO_TRIBUTADO And objTipoTribICMS.iTipo <> ICMS_TIPO_ISENTO _
        And objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO And objTipoTribIPI.iTipo <> IPI_TIPO_ISENTO Then
        sSituacaoTrib = "A"
    
    Else
        
        '***********************************************************************
        ' Avaliacao em relacao ao ICMS
        '***********************************************************************
        
        'Definicao do Tipo B - Tributado de ICMS
        If objTipoTribICMS.iTipo <> ICMS_TIPO_NAO_TRIBUTADO And objTipoTribICMS.iTipo <> ICMS_TIPO_ISENTO Then
            sSituacaoTrib = "B"
        
        'Definicao do Tipo D - Nao Tributado de ICMS
        ElseIf objTipoTribICMS.iTipo = ICMS_TIPO_NAO_TRIBUTADO Then
            sSituacaoTrib = "D"
                
        'Definicao do Tipo F - Isento de ICMS
        ElseIf objTipoTribICMS.iTipo = ICMS_TIPO_ISENTO Then
            sSituacaoTrib = "F"
                
        End If
        
        '***********************************************************************
        ' Avaliacao em relacao ao IPI - Concatenando com a Situacao do ICMS
        '***********************************************************************
                
        'Definicao do Tipo C - Tributado de IPI
        If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO And objTipoTribIPI.iTipo <> IPI_TIPO_ISENTO Then
            sSituacaoTrib = sSituacaoTrib & "C"
        
        'Definicao do Tipo E - Nao Tributado de IPI
        ElseIf objTipoTribIPI.iTipo = IPI_TIPO_NAO_TRIBUTADO Then
            sSituacaoTrib = sSituacaoTrib & "E"
        
        'Definicao do Tipo G - Isento de IPI
        ElseIf objTipoTribIPI.iTipo = IPI_TIPO_ISENTO Then
            sSituacaoTrib = sSituacaoTrib & "G"
        
        End If
        
    End If
    
    Define_SituacaoTributaria_Produto = SUCESSO
    
    Exit Function
    
Erro_Define_SituacaoTributaria_Produto:

    Define_SituacaoTributaria_Produto = gErr
    
    sSituacaoTrib = ""
    
    Select Case gErr
            
        Case 108760, 108762
        
        'Sem dados ICMS
        Case 108761
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS1", gErr, iTipoTribICMS)
            'O Tipo de Tributação ICMS %s não está cadastrado no Banco de Dados.
        
        'Sem dados IPI
        Case 108763
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI1", gErr, iTipoTribIPI)
            'O Tipo de Tributação IPI %s não está cadastrado no Banco de Dados.
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153675)
    
    End Select
    
End Function

Function PIS_Excecao_Pesquisar(ByVal colCateg As Collection, dAliquota As Double) As Long
'retorna o lErro de que nao encontrou excecao
    PIS_Excecao_Pesquisar = 130761
End Function

Function COFINS_Excecao_Pesquisar(ByVal colCateg As Collection, dAliquota As Double) As Long
'retorna o lErro de que nao encontrou excecao
    COFINS_Excecao_Pesquisar = 130767
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TiposTribPISCOFINS_Le_Todos(ByVal colTiposTribPISCOFINS As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTipoTribPISCOFINS As typeTipoTribPISCOFINS
Dim objTipoTribPISCOFINS As ClassTipoTribPISCOFINS

On Error GoTo Erro_TiposTribPISCOFINS_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202829

    'Alocação de espaço no buffer
    tTipoTribPISCOFINS.sDescricao = String(STRING_TIPO_PIS_COFINS_DESCRICAO, 0)

    'Le a tabelaTiposTribPISCOFINS
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, TipoCalculo, VersaoNFe, Entrada, Saida FROM TiposTribPISCOFINS ORDER BY Tipo ", _
                tTipoTribPISCOFINS.iTipo, tTipoTribPISCOFINS.sDescricao, tTipoTribPISCOFINS.iTipoCalculo, tTipoTribPISCOFINS.iVersaoNFE, tTipoTribPISCOFINS.iEntrada, tTipoTribPISCOFINS.iSaida)
    If lErro <> AD_SQL_SUCESSO Then gError 202830

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202831

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTipoTribPISCOFINS = New ClassTipoTribPISCOFINS
    
        objTipoTribPISCOFINS.iTipo = tTipoTribPISCOFINS.iTipo
        objTipoTribPISCOFINS.sDescricao = tTipoTribPISCOFINS.sDescricao
        objTipoTribPISCOFINS.iTipoCalculo = tTipoTribPISCOFINS.iTipoCalculo
        objTipoTribPISCOFINS.iEntrada = tTipoTribPISCOFINS.iEntrada
        objTipoTribPISCOFINS.iSaida = tTipoTribPISCOFINS.iSaida
        objTipoTribPISCOFINS.iVersaoNFE = tTipoTribPISCOFINS.iVersaoNFE
        
        colTiposTribPISCOFINS.Add objTipoTribPISCOFINS

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202832

    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposTribPISCOFINS_Le_Todos = SUCESSO

    Exit Function

Erro_TiposTribPISCOFINS_Le_Todos:

    TiposTribPISCOFINS_Le_Todos = gErr

    Select Case gErr

        Case 202829
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202830 To 202832
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBPISCOFINS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202833)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TributacaoDoc_Le(ByVal objTributacaoDoc As ClassTributacaoDoc) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTributacaoDoc As typeTributacaoDoc

On Error GoTo Erro_TributacaoDoc_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202877
    
    tTributacaoDoc.sNaturezaOpInterna = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaTributacaoDoc
    With tTributacaoDoc
        lErro = Comando_Executar(lComando, "SELECT NaturezaOpInterna, TipoDoc, NumIntDoc, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSBaseManual, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFBaseManual, IRRFAliquota, " & _
                    "IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetido, COFINSRetido, CSLLRetido, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, ISSRetido, ISSRetidoManual, PISCredito, PISCreditoManual, COFINSCredito, COFINSCreditoManual, IPICredito, IPICreditoManual, ICMSCredito, ICMSCreditoManual, ValorINSS, ValorINSSManual, INSSRetido, INSSRetidoManual, INSSValorBase, INSSValorBaseManual, INSSValorDeducoes, INSSValorDeducoesManual, TotTrib, TotTribManual, TotTribTipo, TotTribTipoManual, NaturezaOpManual, " & _
                    "PISValor, COFINSValor, ICMSValorIsento, ISSValorDeducao, ISSValorOutrasRet, ISSValorDescIncond, ISSValorDescCond, IIValor, Destino, DestinoManual, FinalidadeNFe, " & _
                    "FinalidadeNFeManual , IndConsumidorFinal, IndConsumidorFinalManual, IndPresenca, IndPresencaManual, RegTribEspecial, RegTribEspecialManual, DataPrestServico, DataPrestServicoManual,TotTribFederal,TotTribEstadual,TotTribMunicipal,TotTribFederalManual,TotTribEstadualManual,TotTribMunicipalManual,ICMSInterestVlrFCPUFDest,ICMSInterestVlrUFDest,ICMSInterestVlrUFRemet, ICMSVlrFCP, ICMSVlrFCPST, ICMSVlrFCPSTRet, IPIVlrDevolvido " & _
                    " FROM TributacaoDoc WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                    .sNaturezaOpInterna, .iTipoDoc, .lNumIntDoc, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, _
                    .iISSBaseManual, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .iIRRFBaseManual, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, .iPISRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dISSRetido, .iISSRetidoManual, .dPISCredito, _
                    .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iValorINSSManual, .iINSSRetido, .iINSSRetidoManual, .dINSSValorBase, .iINSSValorBaseManual, .dINSSValorDeducoes, .iINSSValorDeducoesManual, .dTotTrib, .iTotTribManual, .iTotTribTipo, .iTotTribTipoManual, .iNaturezaOpManual, _
                    .dPISValor, .dCOFINSValor, .dICMSValorIsento, .dISSValorDeducao, .dISSValorOutrasRet, .dISSValorDescIncond, .dISSValorDescCond, .dIIValor, .iDestino, .iDestinoManual, _
                    .iFinalidadeNFe, .iFinalidadeNFeManual, .iIndConsumidorFinal, .iIndConsumidorFinalManual, .iIndPresenca, .iIndPresencaManual, .iRegTribEspecial, .iRegTribEspecialManual, .dtDataPrestServico, .iDataPrestServicoManual, .dTotTribFed, .dTotTribEst, .dTotTribMunic, .iTotTribFedManual, .iTotTribEstManual, .iTotTribMunicManual, .dICMSInterestVlrFCPUFDest, .dICMSInterestVlrUFDest, .dICMSInterestVlrUFRemet, .dICMSVlrFCP, .dICMSVlrFCPST, .dICMSVlrFCPSTRet, .dIPIVlrDevolvido, _
                    objTributacaoDoc.iTipoDoc, objTributacaoDoc.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 202878

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202879

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    Call objTributacaoDoc.Copia2(tTributacaoDoc)

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TributacaoDoc_Le = SUCESSO

    Exit Function

Erro_TributacaoDoc_Le:

    TributacaoDoc_Le = gErr

    Select Case gErr

        Case 202877
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202878, 202879
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOC", gErr, objTributacaoDoc.iTipoDoc, objTributacaoDoc.lNumIntDoc)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202880)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TributacaoDocItem_Le(ByVal objTributacaoDocItem As ClassTributacaoDocItem) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTributacaoDocItem As typeTributacaoDocItem

On Error GoTo Erro_TributacaoDocItem_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202881

    With tTributacaoDocItem
    
        'Alocação de espaço no buffer
        .sProduto = String(STRING_PRODUTO, 0)
        .sProdutoDescricao = String(STRING_ITEMNF_DESCRICAO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sIPICodProduto = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sCST = String(STRING_CST, 0)
        .sISSQN = String(STRING_ISSQN_CODIGO, 0)
        .sISSCidadeIBGE = String(STRING_CIDADE_CODIBGE, 0)
        .sIPIEnquadramentoClasse = String(STRING_IPI_CLASSE_ENQUADRAMENTO, 0)
        .sIPIEnquadramentoCodigo = String(STRING_IPI_CODIGO_ENQUADRAMENTO, 0)
        .sIPICNPJProdutor = String(STRING_CGC, 0)
        .sIPISeloCodigo = String(STRING_IPI_SELO_CODIGO, 0)
        .sGenero = String(STRING_PRODUTOGENERO_CODIGO, 0)
        .sEAN = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sEANTrib = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sUMTrib = String(STRING_UM_SIGLA, 0)
        .sICMSUFDevidoST = String(STRING_ESTADO, 0)
        .sISSCST = String(STRING_TIPO_ISS_CST, 0)
        .sCSOSN = String(STRING_CSOSN, 0)
        .sNatBCCred = String(STRING_NATBCCRED, 0)
        .sFCI = String(STRING_FCI, 0)
    
        .sISSCodServ = String(STRING_PRODUTO_ISS_CODSERV, 0)
        .sISSMunicIncidImp = String(STRING_CIDADE_CODIBGE, 0)
        .sISSNumProcesso = String(STRING_ISS_NUMPROCESSO, 0)
        .sCEST = String(STRING_MAXIMO, 0)
        .scBenef = String(STRING_MAXIMO, 0)
        .sindEscala = String(STRING_MAXIMO, 0)
        .sCNPJFab = String(STRING_MAXIMO, 0)

        'Le a tabelaTributacaoDocItem
        lErro = Comando_Executar(lComando, "SELECT TipoDoc, NumIntDocItem, NumIntDoc, Item, Produto, ProdutoDescricao, UnidadeMed, Quantidade, PrecoTotal, PrecoUnitario, PrecoUnitarioMoeda, DescontoGrid, NaturezaOp, TipoTributacao, IPITipo, IPICodProduto, IPIBaseCalculo, IPIPercRedBase, IPIAliquota, IPIValor, IPICredito, ICMSTipo, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSCredito, " & _
                    "ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, PISCredito, COFINSCredito, ICMSAliquotaAdicaoDI, ICMSPercRedBaseAdicaoDI, DespImpICMSBase,DespImpICMSValor, DespImpICMSCredito, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, PISValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTAliquotaValor, PISSTQtde, PISSTValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, " & _
                    "COFINSAliquotaValor, COFINSQtde, COFINSValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTAliquotaValor, COFINSSTQtde, COFINSSTValor, CST, ISSQN, ISSBase, ISSAliquota, ISSCidadeIBGE, ISSValor, IIBase, IIDespAduaneira, IIIOF, IIValor, ICMSBaseModalidade, ICMSSubstBaseModalidade, IPIEnquadramentoClasse, IPIEnquadramentoCodigo, IPICNPJProdutor, IPISeloCodigo, IPISeloQtde, IPITipoCalculo, IPIUnidadePadraoQtde, " & _
                    "IPIUnidadePadraoValor, ValorFreteItem, ValorSeguroItem, ValorOutrasDespesasItem, ValorDescontoItem, OrigemMercadoria, ExTIPI, Genero, ProdutoEspecifico, EAN, EANTrib, QtdTrib, UMTrib, ValorUnitTrib, NaturezaOpManual, TipoTributacaoManual, IPITipoManual, IPICodProdutoManual, IPIBaseCalculoManual, IPIPercRedBaseManual, IPIAliquotaManual, IPIValorManual, IPICreditoManual, ICMSTipoManual, ICMSBaseManual, ICMSPercRedBaseManual, ICMSAliquotaManual, ICMSValorManual, " & _
                    "ICMSCreditoManual, ICMSSubstBaseManual, ICMSSubstAliquotaManual, ICMSSubstValorManual, PISCreditoManual, COFINSCreditoManual, ICMSAliquotaAdicaoDIManual, ICMSPercRedBaseAdicaoDIManual, DespImpICMSBaseManual, DespImpICMSValorManual, DespImpICMSCreditoManual, ICMSSubstPercRedBaseManual, ICMSSubstPercMVAManual, PISTipoManual, PISTipoCalculoManual, PISBaseManual, PISAliquotaManual, PISAliquotaValorManual, PISQtdeManual, PISValorManual, PISSTTipoCalculoManual, PISSTBaseManual, PISSTAliquotaManual, PISSTAliquotaValorManual, PISSTQtdeManual, PISSTValorManual, COFINSTipoManual, COFINSTipoCalculoManual, COFINSBaseManual, " & _
                    "COFINSAliquotaManual, COFINSAliquotaValorManual, COFINSQtdeManual, COFINSValorManual, COFINSSTTipoCalculoManual, COFINSSTBaseManual, COFINSSTAliquotaManual, COFINSSTAliquotaValorManual, COFINSSTQtdeManual, COFINSSTValorManual, CSTManual, ISSQNManual, ISSBaseManual, ISSAliquotaManual, ISSCidadeIBGEManual, ISSValorManual, IIBaseManual, IIDespAduaneiraManual, IIIOFManual, IIValorManual, ICMSBaseModalidadeManual, ICMSSubstBaseModalidadeManual, IPIEnquadramentoClasseManual, IPIEnquadramentoCodigoManual, IPICNPJProdutorManual, IPISeloCodigoManual, IPISeloQtdeManual, IPITipoCalculoManual, " & _
                    "IPIUnidadePadraoQtdeManual, IPIUnidadePadraoValorManual, ValorFreteItemManual, ValorSeguroItemManual, ValorOutrasDespesasItemManual, ValorDescontoItemManual, OrigemMercadoriaManual, FCI, FCIManual, ICMSSTCobrAntBase,ICMSSTCobrAntValor,ICMSSTCobrAntBaseManual,ICMSSTCobrAntValorManual, ICMSpercBaseOperacaoPropria,ICMSUFDevidoST,ICMSvBCSTRet,ICMSvICMSSTRet,ICMSvBCSTDest,ICMSvICMSSTDest,ICMSpCredSN,ICMSvCredSN,ISSCST,ICMSValorIsento,ICMSMotivo,ICMSSimplesTipo,ICMSpercBaseOperacaoPropriaManual,ICMSUFDevidoSTManual," & _
                    "ICMSvBCSTRetManual,ICMSvICMSSTRetManual,ICMSvBCSTDestManual,ICMSvICMSSTDestManual,ICMSpCredSNManual,ICMSvCredSNManual, ICMSValorIsentoManual, ICMSMotivoManual, ICMSSimplesTipoManual, ISSTipo, ISSTipoManual, RegimeTributario, CSOSN, NatBCCred, NatBCCredManual, TotTrib, TotTribManual, TotTribTipo, TotTribTipoManual, ICMS51ValorOp, ICMSPercDifer, ICMSValorDif, ICMS51ValorOpManual, ICMSPercDiferManual, ICMSValorDifManual, ISSValorDeducao, ISSValorOutrasRet, ISSValorDescIncond, ISSValorDescCond, ISSValorRet, ISSIndExigibilidade, ISSCodServ, ISSMunicIncidImp, ISSCodPais, ISSNumProcesso, ISSIndIncentivo, ISSValorDeducaoManual, ISSValorOutrasRetManual, " & _
                    "ISSValorDescIncondManual , ISSValorDescCondManual, ISSValorRetManual, ISSIndExigibilidadeManual, ISSCodServManual, ISSMunicIncidImpManual, ISSCodPaisManual, ISSNumProcessoManual, ISSIndIncentivoManual,TotTribAliqFederal,TotTribAliqEstadual,TotTribAliqMunicipal,TotTribAliqFederalManual,TotTribAliqEstadualManual,TotTribAliqMunicipalManual, " & _
                    "CEST,ICMSInterestBCUFDest,ICMSInterestBCUFDestManual,ICMSInterestPercFCPUFDest,ICMSInterestPercFCPUFDestManual,ICMSInterestAliqUFDest,ICMSInterestAliqUFDestManual,ICMSInterestPercPartilha,ICMSInterestPercPartilhaManual,ICMSInterestVlrFCPUFDest,ICMSInterestVlrFCPUFDestManual,ICMSInterestVlrUFDest,ICMSInterestVlrUFDestManual,ICMSInterestVlrUFRemet,ICMSInterestVlrUFRemetManual,ICMSInterestAliq,ICMSInterestAliqManual, " & _
                    "ICMSInterestBCFCPUFDest,ICMSInterestBCFCPUFDestManual,ICMSvBCFCPST,ICMSvBCFCPSTManual,ICMSpFCPST,ICMSpFCPSTManual,ICMSvFCPST,ICMSvFCPSTManual,ICMSvBCFCPSTRet,ICMSvBCFCPSTRetManual,ICMSpFCPSTRet,ICMSpFCPSTRetManual,ICMSvFCPSTRet,ICMSvFCPSTRetManual,ICMSSTCobrAntAliquota,ICMSSTCobrAntAliquotaManual,ICMSvBCFCP,ICMSvBCFCPManual,ICMSpFCP,ICMSpFCPManual,ICMSvFCP,ICMSvFCPManual,cBenef,indEscala,CNPJFab,IPIVlrDevolvido,IPIVlrDevolvidoManual,pDevol,pDevolManual,ICMSSTBaseDupla,ICMSSTBaseDuplaManual,ICMSInterestBaseDupla,ICMSInterestBaseDuplaManual, cBenefManual FROM TributacaoDocItem WHERE TipoDoc= ?  AND NumIntDocItem= ? ", _
                    .iTipoDoc, .lNumIntDocItem, .lNumIntDoc, .iItem, .sProduto, .sProdutoDescricao, .sUnidadeMed, .dQuantidade, .dPrecoTotal, .dPrecoUnitario, .dPrecoUnitarioMoeda, .dDescontoGrid, .sNaturezaOp, .iTipoTributacao, _
                    .iIPITipo, .sIPICodProduto, .dIPIBaseCalculo, .dIPIPercRedBase, .dIPIAliquota, .dIPIValor, .dIPICredito, .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstAliquota, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito, .dICMSAliquotaAdicaoDI, .dICMSPercRedBaseAdicaoDI, .dDespImpICMSBase, .dDespImpICMSValor, .dDespImpICMSCredito, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .iPISTipo, .iPISTipoCalculo, .dPISBase, .dPISAliquota, .dPISAliquotaValor, .dPISQtde, _
                    .dPISValor, .iPISSTTipoCalculo, .dPISSTBase, .dPISSTAliquota, .dPISSTAliquotaValor, .dPISSTQtde, .dPISSTValor, .iCOFINSTipo, .iCOFINSTipoCalculo, .dCOFINSBase, .dCOFINSAliquota, .dCOFINSAliquotaValor, .dCOFINSQtde, .dCOFINSValor, .iCOFINSSTTipoCalculo, .dCOFINSSTBase, .dCOFINSSTAliquota, .dCOFINSSTAliquotaValor, .dCOFINSSTQtde, .dCOFINSSTValor, .sCST, .sISSQN, .dISSBase, .dISSAliquota, .sISSCidadeIBGE, .dISSValor, .dIIBase, .dIIDespAduaneira, .dIIIOF, .dIIValor, .iICMSBaseModalidade, .iICMSSubstBaseModalidade, .sIPIEnquadramentoClasse, .sIPIEnquadramentoCodigo, .sIPICNPJProdutor, .sIPISeloCodigo, .lIPISeloQtde, .iIPITipoCalculo, .dIPIUnidadePadraoQtde, .dIPIUnidadePadraoValor, _
                    .dValorFreteItem, .dValorSeguroItem, .dValorOutrasDespesasItem, .dValorDescontoItem, .iOrigemMercadoria, .iExTIPI, .sGenero, .iProdutoEspecifico, .sEAN, .sEANTrib, .dQtdTrib, .sUMTrib, .dValorUnitTrib, .iNaturezaOpManual, .iTipoTributacaoManual, .iIPITipoManual, .iIPICodProdutoManual, .iIPIBaseCalculoManual, .iIPIPercRedBaseManual, .iIPIAliquotaManual, .iIPIValorManual, .iIPICreditoManual, .iICMSTipoManual, .iICMSBaseManual, .iICMSPercRedBaseManual, .iICMSAliquotaManual, .iICMSValorManual, .iICMSCreditoManual, .iICMSSubstBaseManual, .iICMSSubstAliquotaManual, .iICMSSubstValorManual, .iPISCreditoManual, .iCOFINSCreditoManual, .iICMSAliquotaAdicaoDIManual, .iICMSPercRedBaseAdicaoDIManual, .iDespImpICMSBaseManual, .iDespImpICMSValorManual, .iDespImpICMSCreditoManual, .iICMSSubstPercRedBaseManual, .iICMSSubstPercMVAManual, .iPISTipoManual, _
                    .iPISTipoCalculoManual, .iPISBaseManual, .iPISAliquotaManual, .iPISAliquotaValorManual, .iPISQtdeManual, .iPISValorManual, .iPISSTTipoCalculoManual, .iPISSTBaseManual, .iPISSTAliquotaManual, .iPISSTAliquotaValorManual, .iPISSTQtdeManual, .iPISSTValorManual, .iCOFINSTipoManual, .iCOFINSTipoCalculoManual, .iCOFINSBaseManual, .iCOFINSAliquotaManual, .iCOFINSAliquotaValorManual, .iCOFINSQtdeManual, .iCOFINSValorManual, .iCOFINSSTTipoCalculoManual, _
                    .iCOFINSSTBaseManual, .iCOFINSSTAliquotaManual, .iCOFINSSTAliquotaValorManual, .iCOFINSSTQtdeManual, .iCOFINSSTValorManual, .iCSTManual, .iISSQNManual, .iISSBaseManual, .iISSAliquotaManual, .iISSCidadeIBGEManual, .iISSValorManual, .iIIBaseManual, .iIIDespAduaneiraManual, .iIIIOFManual, .iIIValorManual, .iICMSBaseModalidadeManual, .iICMSSubstBaseModalidadeManual, .iIPIEnquadramentoClasseManual, .iIPIEnquadramentoCodigoManual, .iIPICNPJProdutorManual, _
                    .iIPISeloCodigoManual, .iIPISeloQtdeManual, .iIPITipoCalculoManual, .iIPIUnidadePadraoQtdeManual, .iIPIUnidadePadraoValorManual, .iValorFreteItemManual, .iValorSeguroItemManual, .iValorOutrasDespesasItemManual, .iValorDescontoItemManual, .iOrigemMercadoriaManual, .sFCI, .iFCIManual, .dICMSSTCobrAntBase, .dICMSSTCobrAntValor, .iICMSSTCobrAntBaseManual, .iICMSSTCobrAntValorManual, _
                    .dICMSpercBaseOperacaoPropria, .sICMSUFDevidoST, .dICMSvBCSTRet, .dICMSvICMSSTRet, .dICMSvBCSTDest, .dICMSvICMSSTDest, .dICMSpCredSN, .dICMSvCredSN, .sISSCST, .dICMSValorIsento, .iICMSMotivo, .iICMSSimplesTipo, .iICMSpercBaseOperacaoPropriaManual, .iICMSUFDevidoSTManual, .iICMSvBCSTRetManual, .iICMSvICMSSTRetManual, .iICMSvBCSTDestManual, .iICMSvICMSSTDestManual, .iICMSpCredSNManual, .iICMSvCredSNManual, .iICMSValorIsentoManual, .iICMSMotivoManual, .iICMSSimplesTipoManual, .iISSTipo, .iISSTipoManual, .iRegimeTributario, .sCSOSN, .sNatBCCred, .iNatBCCredManual, .dTotTrib, .iTotTribManual, .iTotTribTipo, .iTotTribTipoManual, _
                    .dICMS51ValorOp, .dICMSPercDifer, .dICMSValorDif, .iICMS51ValorOpManual, .iICMSPercDiferManual, .iICMSValorDifManual, .dISSValorDeducao, .dISSValorOutrasRet, .dISSValorDescIncond, .dISSValorDescCond, .dISSValorRet, .iISSIndExigibilidade, .sISSCodServ, .sISSMunicIncidImp, .iISSCodPais, .sISSNumProcesso, .iISSIndIncentivo, .iISSValorDeducaoManual, .iISSValorOutrasRetManual, .iISSValorDescIncondManual, _
                    .iISSValorDescCondManual, .iISSValorRetManual, .iISSIndExigibilidadeManual, .iISSCodServManual, .iISSMunicIncidImpManual, .iISSCodPaisManual, .iISSNumProcessoManual, .iISSIndIncentivoManual, .dTotTribAliqFed, .dTotTribAliqEst, .dTotTribAliqMunic, .iTotTribAliqFedManual, .iTotTribAliqEstManual, .iTotTribAliqMunicManual, _
                    .sCEST, .dICMSInterestBCUFDest, .iICMSInterestBCUFDestManual, .dICMSInterestPercFCPUFDest, .iICMSInterestPercFCPUFDestManual, .dICMSInterestAliqUFDest, .iICMSInterestAliqUFDestManual, .dICMSInterestPercPartilha, .iICMSInterestPercPartilhaManual, .dICMSInterestVlrFCPUFDest, .iICMSInterestVlrFCPUFDestManual, .dICMSInterestVlrUFDest, .iICMSInterestVlrUFDestManual, .dICMSInterestVlrUFRemet, .iICMSInterestVlrUFRemetManual, .dICMSInterestAliq, .iICMSInterestAliqManual, _
                    .dICMSInterestBCFCPUFDest, .iICMSInterestBCFCPUFDestManual, .dICMSvBCFCPST, .iICMSvBCFCPSTManual, .dICMSpFCPST, .iICMSpFCPSTManual, .dICMSvFCPST, .iICMSvFCPSTManual, .dICMSvBCFCPSTRet, .iICMSvBCFCPSTRetManual, .dICMSpFCPSTRet, .iICMSpFCPSTRetManual, .dICMSvFCPSTRet, .iICMSvFCPSTRetManual, .dICMSSTCobrAntAliquota, .iICMSSTCobrAntAliquotaManual, .dICMSvBCFCP, .iICMSvBCFCPManual, .dICMSpFCP, .iICMSpFCPManual, .dICMSvFCP, .iICMSvFCPManual, .scBenef, .sindEscala, .sCNPJFab, .dIPIVlrDevolvido, .iIPIVlrDevolvidoManual, .dpDevol, .ipDevolManual, .iICMSSTBaseDupla, .iICMSSTBaseDuplaManual, .iICMSInterestBaseDupla, .iICMSInterestBaseDuplaManual, .icBenefManual, objTributacaoDocItem.iTipoDoc, objTributacaoDocItem.lNumIntDocItem)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 202882

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202883

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    'Força ser automático porque senão dá erro ao reutilizar um documento
    'mexendo na informação a nível de nota
    tTributacaoDocItem.iValorDescontoItemManual = VAR_PREENCH_AUTOMATICO
    tTributacaoDocItem.iValorFreteItemManual = VAR_PREENCH_AUTOMATICO
    tTributacaoDocItem.iValorOutrasDespesasItemManual = VAR_PREENCH_AUTOMATICO
    tTributacaoDocItem.iValorSeguroItemManual = VAR_PREENCH_AUTOMATICO

    Call objTributacaoDocItem.Copia2(tTributacaoDocItem)

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TributacaoDocItem_Le = SUCESSO

    Exit Function

Erro_TributacaoDocItem_Le:

    TributacaoDocItem_Le = gErr

    Select Case gErr

        Case 202881
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202882, 202883
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOCITEM", gErr, objTributacaoDocItem.iTipoDoc, objTributacaoDocItem.lNumIntDocItem)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202884)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PISCOFINSExcecoes_Le(ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPISCOFINSExcecao As typePISCOFINSExcecao

On Error GoTo Erro_PISCOFINSExcecoes_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205473

    'Alocação de espaço no buffer
    tPISCOFINSExcecao.sFundamentacao = String(STRING_EXCECAO_FUNDAMENTACAO, 0)

    'Le a tabelaPISCOFINSExcecoes
    lErro = Comando_Executar(lComando, "SELECT Tipo, TipoPIS, TipoCOFINS, PISTipoCalculo, COFINSTipoCalculo, AliquotaPisRS, AliquotaPisPerc, AliquotaCofinsRS, AliquotaCofinsPerc, " & _
                "Fundamentacao, Prioridade, TipoCliForn, TipoPISE, TipoCOFINSE FROM PISCOFINSExcecoes WHERE CategoriaCliente= ?  AND CategoriaClienteItem= ?  AND CategoriaProduto= ?  AND CategoriaProdutoItem= ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem = ?", _
                tPISCOFINSExcecao.iTipo, tPISCOFINSExcecao.iTipoPIS, tPISCOFINSExcecao.iTipoCOFINS, tPISCOFINSExcecao.iPISTipoCalculo, tPISCOFINSExcecao.iCOFINSTipoCalculo, _
                tPISCOFINSExcecao.dAliquotaPisRS, tPISCOFINSExcecao.dAliquotaPisPerc, tPISCOFINSExcecao.dAliquotaCofinsRS, tPISCOFINSExcecao.dAliquotaCofinsPerc, tPISCOFINSExcecao.sFundamentacao, _
                tPISCOFINSExcecao.iPrioridade, tPISCOFINSExcecao.iTipoCliForn, tPISCOFINSExcecao.iTipoPISE, tPISCOFINSExcecao.iTipoCOFINSE, _
                objPISCOFINSExcecao.sCategoriaCliente, objPISCOFINSExcecao.sCategoriaClienteItem, objPISCOFINSExcecao.sCategoriaProduto, objPISCOFINSExcecao.sCategoriaProdutoItem, objPISCOFINSExcecao.sCategoriaFornecedor, objPISCOFINSExcecao.sCategoriaFornecedorItem)
    If lErro <> AD_SQL_SUCESSO Then gError 205474

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205475

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPISCOFINSExcecao.iTipo = tPISCOFINSExcecao.iTipo
    objPISCOFINSExcecao.iTipoPIS = tPISCOFINSExcecao.iTipoPIS
    objPISCOFINSExcecao.iTipoCOFINS = tPISCOFINSExcecao.iTipoCOFINS
    objPISCOFINSExcecao.iPISTipoCalculo = tPISCOFINSExcecao.iPISTipoCalculo
    objPISCOFINSExcecao.iCOFINSTipoCalculo = tPISCOFINSExcecao.iCOFINSTipoCalculo
    objPISCOFINSExcecao.dAliquotaPisRS = tPISCOFINSExcecao.dAliquotaPisRS
    objPISCOFINSExcecao.dAliquotaPisPerc = tPISCOFINSExcecao.dAliquotaPisPerc
    objPISCOFINSExcecao.dAliquotaCofinsRS = tPISCOFINSExcecao.dAliquotaCofinsRS
    objPISCOFINSExcecao.dAliquotaCofinsPerc = tPISCOFINSExcecao.dAliquotaCofinsPerc
    objPISCOFINSExcecao.sFundamentacao = tPISCOFINSExcecao.sFundamentacao
    objPISCOFINSExcecao.iPrioridade = tPISCOFINSExcecao.iPrioridade
    objPISCOFINSExcecao.iTipoCliForn = tPISCOFINSExcecao.iTipoCliForn
    objPISCOFINSExcecao.iTipoPISE = tPISCOFINSExcecao.iTipoPISE
    objPISCOFINSExcecao.iTipoCOFINSE = tPISCOFINSExcecao.iTipoCOFINSE

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PISCOFINSExcecoes_Le = SUCESSO

    Exit Function

Erro_PISCOFINSExcecoes_Le:

    PISCOFINSExcecoes_Le = gErr

    Select Case gErr

        Case 205473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205474, 205475
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PISCOFINSEXCECOES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205476)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PISCOFINSExcecao_Le_Cat(ByVal lCliente As Long, ByVal iFilialCliente As Integer, ByVal sProduto As String, ByVal iTipo As Integer, ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPISCOFINSExcecao As typePISCOFINSExcecao
Dim sSQL As String

On Error GoTo Erro_PISCOFINSExcecao_Le_Cat

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205473
    
    sSQL = "SELECT Tipo, TipoPIS, TipoCOFINS, PISTipoCalculo, COFINSTipoCalculo, AliquotaPisRS, AliquotaPisPerc, AliquotaCofinsRS, AliquotaCofinsPerc, Fundamentacao, Prioridade, TipoPISE, TipoCOFINSE "
    sSQL = sSQL & "FROM PISCOFINSExcecoes AS E "
    sSQL = sSQL & "WHERE TipoCliForn = 0 AND ((CategoriaCliente = '' OR EXISTS ( "
    sSQL = sSQL & "SELECT Cliente FROM FilialClienteCategorias AS F "
    sSQL = sSQL & "WHERE F.Cliente = ? AND F.Filial = ? AND F.Categoria = E.CategoriaCliente AND F.Item = E.CategoriaClienteItem "
    sSQL = sSQL & ")) AND "
    sSQL = sSQL & "(CategoriaProduto = '' OR EXISTS ( "
    sSQL = sSQL & "SELECT Produto FROM ProdutoCategoria AS P "
    sSQL = sSQL & "WHERE P.Produto = ? AND P.Categoria = E.CategoriaProduto AND P.Item = E.CategoriaProdutoItem "
    sSQL = sSQL & ")) AND Tipo IN (?,?)) "
    sSQL = sSQL & "ORDER BY Prioridade "

    'Alocação de espaço no buffer
    tPISCOFINSExcecao.sFundamentacao = String(STRING_EXCECAO_FUNDAMENTACAO, 0)

    'Le a tabelaPISCOFINSExcecoes
    lErro = Comando_Executar(lComando, sSQL, _
                tPISCOFINSExcecao.iTipo, tPISCOFINSExcecao.iTipoPIS, tPISCOFINSExcecao.iTipoCOFINS, tPISCOFINSExcecao.iPISTipoCalculo, tPISCOFINSExcecao.iCOFINSTipoCalculo, _
                tPISCOFINSExcecao.dAliquotaPisRS, tPISCOFINSExcecao.dAliquotaPisPerc, tPISCOFINSExcecao.dAliquotaCofinsRS, tPISCOFINSExcecao.dAliquotaCofinsPerc, tPISCOFINSExcecao.sFundamentacao, _
                tPISCOFINSExcecao.iPrioridade, tPISCOFINSExcecao.iTipoPISE, tPISCOFINSExcecao.iTipoCOFINSE, _
                lCliente, iFilialCliente, sProduto, iTipo, EXCECAO_PIS_COFINS_TIPO_AMBOS)
    If lErro <> AD_SQL_SUCESSO Then gError 205474

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205475

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPISCOFINSExcecao.iTipo = tPISCOFINSExcecao.iTipo
    objPISCOFINSExcecao.iTipoPIS = tPISCOFINSExcecao.iTipoPIS
    objPISCOFINSExcecao.iTipoCOFINS = tPISCOFINSExcecao.iTipoCOFINS
    objPISCOFINSExcecao.iPISTipoCalculo = tPISCOFINSExcecao.iPISTipoCalculo
    objPISCOFINSExcecao.iCOFINSTipoCalculo = tPISCOFINSExcecao.iCOFINSTipoCalculo
    objPISCOFINSExcecao.dAliquotaPisRS = tPISCOFINSExcecao.dAliquotaPisRS
    objPISCOFINSExcecao.dAliquotaPisPerc = tPISCOFINSExcecao.dAliquotaPisPerc
    objPISCOFINSExcecao.dAliquotaCofinsRS = tPISCOFINSExcecao.dAliquotaCofinsRS
    objPISCOFINSExcecao.dAliquotaCofinsPerc = tPISCOFINSExcecao.dAliquotaCofinsPerc
    objPISCOFINSExcecao.sFundamentacao = tPISCOFINSExcecao.sFundamentacao
    objPISCOFINSExcecao.iPrioridade = tPISCOFINSExcecao.iPrioridade
    objPISCOFINSExcecao.iTipoPISE = tPISCOFINSExcecao.iTipoPISE
    objPISCOFINSExcecao.iTipoCOFINSE = tPISCOFINSExcecao.iTipoCOFINSE

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PISCOFINSExcecao_Le_Cat = SUCESSO

    Exit Function

Erro_PISCOFINSExcecao_Le_Cat:

    PISCOFINSExcecao_Le_Cat = gErr

    Select Case gErr

        Case 205473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205474, 205475
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PISCOFINSEXCECOES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205476)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PISCOFINSExcecao_Le_CatForn(ByVal lFornecedor As Long, ByVal iFilialFornecedor As Integer, ByVal sProduto As String, ByVal iTipo As Integer, ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPISCOFINSExcecao As typePISCOFINSExcecao
Dim sSQL As String

On Error GoTo Erro_PISCOFINSExcecao_Le_CatForn

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205473
    
    sSQL = "SELECT Tipo, TipoPIS, TipoCOFINS, PISTipoCalculo, COFINSTipoCalculo, AliquotaPisRS, AliquotaPisPerc, AliquotaCofinsRS, AliquotaCofinsPerc, Fundamentacao, Prioridade, TipoPISE, TipoCOFINSE "
    sSQL = sSQL & "FROM PISCOFINSExcecoes AS E "
    sSQL = sSQL & "WHERE TipoCliForn = 1 AND ((CategoriaFornecedor = '' OR EXISTS ( "
    sSQL = sSQL & "SELECT Fornecedor FROM FilialFornecedorCategorias AS F "
    sSQL = sSQL & "WHERE F.Fornecedor = ? AND F.Filial = ? AND F.Categoria = E.CategoriaFornecedor AND F.Item = E.CategoriaFornecedorItem "
    sSQL = sSQL & ")) AND "
    sSQL = sSQL & "(CategoriaProduto = '' OR EXISTS ( "
    sSQL = sSQL & "SELECT Produto FROM ProdutoCategoria AS P "
    sSQL = sSQL & "WHERE P.Produto = ? AND P.Categoria = E.CategoriaProduto AND P.Item = E.CategoriaProdutoItem "
    sSQL = sSQL & ")) AND Tipo IN (?,?)) "
    sSQL = sSQL & "ORDER BY Prioridade "

    'Alocação de espaço no buffer
    tPISCOFINSExcecao.sFundamentacao = String(STRING_EXCECAO_FUNDAMENTACAO, 0)

    'Le a tabelaPISCOFINSExcecoes
    lErro = Comando_Executar(lComando, sSQL, _
                tPISCOFINSExcecao.iTipo, tPISCOFINSExcecao.iTipoPIS, tPISCOFINSExcecao.iTipoCOFINS, tPISCOFINSExcecao.iPISTipoCalculo, tPISCOFINSExcecao.iCOFINSTipoCalculo, _
                tPISCOFINSExcecao.dAliquotaPisRS, tPISCOFINSExcecao.dAliquotaPisPerc, tPISCOFINSExcecao.dAliquotaCofinsRS, tPISCOFINSExcecao.dAliquotaCofinsPerc, tPISCOFINSExcecao.sFundamentacao, _
                tPISCOFINSExcecao.iPrioridade, tPISCOFINSExcecao.iTipoPISE, tPISCOFINSExcecao.iTipoCOFINSE, _
                lFornecedor, iFilialFornecedor, sProduto, iTipo, EXCECAO_PIS_COFINS_TIPO_AMBOS)
    If lErro <> AD_SQL_SUCESSO Then gError 205474

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205475

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPISCOFINSExcecao.iTipo = tPISCOFINSExcecao.iTipo
    objPISCOFINSExcecao.iTipoPIS = tPISCOFINSExcecao.iTipoPIS
    objPISCOFINSExcecao.iTipoCOFINS = tPISCOFINSExcecao.iTipoCOFINS
    objPISCOFINSExcecao.iPISTipoCalculo = tPISCOFINSExcecao.iPISTipoCalculo
    objPISCOFINSExcecao.iCOFINSTipoCalculo = tPISCOFINSExcecao.iCOFINSTipoCalculo
    objPISCOFINSExcecao.dAliquotaPisRS = tPISCOFINSExcecao.dAliquotaPisRS
    objPISCOFINSExcecao.dAliquotaPisPerc = tPISCOFINSExcecao.dAliquotaPisPerc
    objPISCOFINSExcecao.dAliquotaCofinsRS = tPISCOFINSExcecao.dAliquotaCofinsRS
    objPISCOFINSExcecao.dAliquotaCofinsPerc = tPISCOFINSExcecao.dAliquotaCofinsPerc
    objPISCOFINSExcecao.sFundamentacao = tPISCOFINSExcecao.sFundamentacao
    objPISCOFINSExcecao.iPrioridade = tPISCOFINSExcecao.iPrioridade
    objPISCOFINSExcecao.iTipoPISE = tPISCOFINSExcecao.iTipoPISE
    objPISCOFINSExcecao.iTipoCOFINSE = tPISCOFINSExcecao.iTipoCOFINSE

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PISCOFINSExcecao_Le_CatForn = SUCESSO

    Exit Function

Erro_PISCOFINSExcecao_Le_CatForn:

    PISCOFINSExcecao_Le_CatForn = gErr

    Select Case gErr

        Case 205473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205474, 205475
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PISCOFINSEXCECOES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205476)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposTribISS_Le_Todos(ByVal colTiposTribISS As Collection) As Long

'preenche colTiposTribISS com objetos ClassTipoTribISS a partir da tabela TiposTribISS
Dim lErro As Long, lComando As Long
Dim objTipoTribISS As ClassTipoTribISS
Dim iComRetencao As Integer, iTipo As Integer, sDescricao As String
Dim iVersaoNFE As Integer, sCST As String

On Error GoTo Erro_TiposTribISS_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209162
    
    sDescricao = String(STRING_TIPO_ISS_DESCRICAO, 0)
    sCST = String(STRING_TIPO_ISS_CST, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, ComRetencao, VersaoNFe, TipoTribCST FROM TiposTribISS ORDER BY Tipo", _
        iTipo, sDescricao, iComRetencao, iVersaoNFE, sCST)
    If lErro <> AD_SQL_SUCESSO Then gError 209163
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209164
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objTipoTribISS = New ClassTipoTribISS

        objTipoTribISS.iComRetencao = iComRetencao
        objTipoTribISS.iTipo = iTipo
        objTipoTribISS.sDescricao = sDescricao
        objTipoTribISS.iVersaoNFE = iVersaoNFE
        objTipoTribISS.sTipoTribCST = sCST
        
        colTiposTribISS.Add objTipoTribISS
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209165
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    TiposTribISS_Le_Todos = SUCESSO

    Exit Function

Erro_TiposTribISS_Le_Todos:

    TiposTribISS_Le_Todos = gErr

    Select Case gErr

        Case 209162
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209163 To 209165
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBISS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209166)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TiposTribICMSSimples_Le_Todos(ByVal colTiposTribICMS As Collection) As Long
'preenche colTiposTribICMS com objetos ClassTipoTribICMS a partir da tabela TiposTribICMS

Dim lErro As Long, lComando As Long
Dim objTipoTribICMS As ClassTipoTribICMSSimples
Dim iPermiteCredito As Integer, iTipo As Integer, sDescricao As String
Dim iPermiteST As Integer, iCSOSN As Integer, iVersaoNFE As Integer
Dim iPermiteAliquota As Integer, iPermiteSTRetAnt As Integer

On Error GoTo Erro_TiposTribICMSSimples_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209167
    
    sDescricao = String(STRING_TIPO_ICMS_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, PermiteCredito, PermiteST, CSOSN, VersaoNFE, PermiteAliquota, PermiteSTRetAnt FROM TiposTribICMSSimples ORDER BY CSOSN", _
        iTipo, sDescricao, iPermiteCredito, iPermiteST, iCSOSN, iVersaoNFE, iPermiteAliquota, iPermiteSTRetAnt)
    If lErro <> AD_SQL_SUCESSO Then gError 209168
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209169
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objTipoTribICMS = New ClassTipoTribICMSSimples

        objTipoTribICMS.iPermiteCredito = iPermiteCredito
        objTipoTribICMS.iTipo = iTipo
        objTipoTribICMS.sDescricao = sDescricao
        objTipoTribICMS.iPermiteST = iPermiteST
        objTipoTribICMS.iCSOSN = iCSOSN
        objTipoTribICMS.iVersaoNFE = iVersaoNFE
        objTipoTribICMS.iPermiteAliquota = iPermiteAliquota
        objTipoTribICMS.iPermiteSTRetAnt = iPermiteSTRetAnt
        
        colTiposTribICMS.Add objTipoTribICMS
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209170
    
    Loop
       
    lErro = Comando_Fechar(lComando)
    
    TiposTribICMSSimples_Le_Todos = SUCESSO

    Exit Function

Erro_TiposTribICMSSimples_Le_Todos:

    TiposTribICMSSimples_Le_Todos = gErr

    Select Case gErr

        Case 209167
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209168 To 209170
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBICMS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209171)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoTribICMSSimples_Le(ByVal objTipoTribICMS As ClassTipoTribICMSSimples) As Long
'preenche colTiposTribICMS com objetos ClassTipoTribICMS a partir da tabela TiposTribICMS

Dim lErro As Long, lComando As Long
Dim iPermiteCredito As Integer, iTipo As Integer, sDescricao As String
Dim iPermiteST As Integer, iCSOSN As Integer, iVersaoNFE As Integer
Dim iPermiteAliquota As Integer, iPermiteSTRetAnt As Integer

On Error GoTo Erro_TipoTribICMSSimples_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209167
    
    sDescricao = String(STRING_TIPO_ICMS_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, PermiteCredito, PermiteST, CSOSN, VersaoNFE, PermiteAliquota, PermiteSTRetAnt FROM TiposTribICMSSimples WHERE Tipo = ?", _
        iTipo, sDescricao, iPermiteCredito, iPermiteST, iCSOSN, iVersaoNFE, iPermiteAliquota, iPermiteSTRetAnt, objTipoTribICMS.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 209168
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209169
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTipoTribICMS.iPermiteCredito = iPermiteCredito
    objTipoTribICMS.iTipo = iTipo
    objTipoTribICMS.sDescricao = sDescricao
    objTipoTribICMS.iPermiteST = iPermiteST
    objTipoTribICMS.iCSOSN = iCSOSN
    objTipoTribICMS.iVersaoNFE = iVersaoNFE
    objTipoTribICMS.iPermiteAliquota = iPermiteAliquota
    objTipoTribICMS.iPermiteSTRetAnt = iPermiteSTRetAnt
      
    Call Comando_Fechar(lComando)
    
    TipoTribICMSSimples_Le = SUCESSO

    Exit Function

Erro_TipoTribICMSSimples_Le:

    TipoTribICMSSimples_Le = gErr

    Select Case gErr

        Case 209167
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209168 To 209169
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBICMS", gErr)
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209171)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function IBPTax_ObterPerc(ByVal iFilialEmpresa As Integer, ByVal dtDataEmissao As Date, ByVal objProduto As ClassProduto, ByVal iOrigemMercadoria As Integer, dPerc As Double, ByVal sUF As String, dTotTribAliqFed As Double, dTotTribAliqEst As Double, dTotTribAliqMunic As Double) As Long

Dim lErro As Long, alComando(0 To 4) As Long, lCodVersao As Long, sCampo As String
Dim iIndice As Integer

On Error GoTo Erro_IBPTax_ObterPerc

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201521
    Next
    
    lErro = Comando_Executar(alComando(0), "SELECT MAX(CodVersao) FROM IBPTax WHERE Validade <= ?", lCodVersao, dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then gError 201522
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201523
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_ITEM_NAO_CADASTRADO
    
    If lCodVersao < 9 Then sUF = ""
    
    Select Case iOrigemMercadoria
    
        Case 0, 3, 4, 5
            sCampo = "AliqNac, AliqNacFed"
            
        Case Else
            sCampo = "AliqImp, AliqImpFed"
            
    End Select
    
    sCampo = sCampo & ", AliqEstadual, AliqMunicipal"
        
    lErro = Comando_Executar(alComando(1), "SELECT " & sCampo & " FROM IBPTaxAliq WHERE CodVersao = ? AND NCM = ? AND ExTIPI = ? AND Tabela = 0 AND UF = ?", dPerc, dTotTribAliqFed, dTotTribAliqEst, dTotTribAliqMunic, lCodVersao, objProduto.sIPICodigo, objProduto.sIPICodDIPI, sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 201524
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
    If lErro <> AD_SQL_SUCESSO Then
        
        lErro = Comando_Executar(alComando(4), "SELECT " & sCampo & " FROM IBPTaxAliq WHERE CodVersao = ? AND ? LIKE NCM AND ExTIPI = ? AND Tabela = 0 AND UF = ? ORDER BY NCM DESC", dPerc, dTotTribAliqFed, dTotTribAliqEst, dTotTribAliqMunic, lCodVersao, objProduto.sIPICodigo, objProduto.sIPICodDIPI, sUF)
        If lErro <> AD_SQL_SUCESSO Then gError 201524
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
        If lErro <> AD_SQL_SUCESSO Then
        
            lErro = Comando_Executar(alComando(2), "SELECT " & sCampo & " FROM IBPTaxAliq WHERE CodVersao = ? AND NCM = ? AND ExTIPI = '' AND Tabela = 1 AND UF=?", dPerc, dTotTribAliqFed, dTotTribAliqEst, dTotTribAliqMunic, lCodVersao, objProduto.sNBS, sUF)
            If lErro <> AD_SQL_SUCESSO Then gError 201524
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
            If lErro <> AD_SQL_SUCESSO Then
            
                lErro = Comando_Executar(alComando(3), "SELECT " & sCampo & " FROM IBPTaxAliq WHERE CodVersao = ? AND NCM = ? AND ExTIPI = '' AND Tabela = 2 AND UF=?", dPerc, dTotTribAliqFed, dTotTribAliqEst, dTotTribAliqMunic, lCodVersao, objProduto.sISSQN, sUF)
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
                If lErro <> AD_SQL_SUCESSO Then gError ERRO_ITEM_NAO_CADASTRADO
        
            End If
        
        End If
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    IBPTax_ObterPerc = SUCESSO
    
    Exit Function
    
Erro_IBPTax_ObterPerc:

    IBPTax_ObterPerc = gErr

    Select Case gErr

        Case 201521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case ERRO_ITEM_NAO_CADASTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201520)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function IPICodEnquadramento_Le(ByVal objIPICodEnquadramento As ClassIPICodEnquadramento) As Long

Dim lErro As Long
Dim lComando As Long
Dim tIPICodEnquadramento As typeIPICodEnquadramento

On Error GoTo Erro_IPICodEnquadramento_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 216105

    'Alocação de espaço no buffer
    tIPICodEnquadramento.sCodigo = String(STRING_MAXIMO, 0)
    tIPICodEnquadramento.sGrupoCST = String(STRING_MAXIMO, 0)
    tIPICodEnquadramento.sDescCompleta = String(STRING_IPICODENQ_DESC, 0)

    'Le a tabelaIPICodEnquadramento
    lErro = Comando_Executar(lComando, "SELECT Codigo, GrupoCST, DescCompleta, TipoIPI FROM IPICodEnquadramento WHERE Codigo= ? ", _
                tIPICodEnquadramento.sCodigo, tIPICodEnquadramento.sGrupoCST, tIPICodEnquadramento.sDescCompleta, tIPICodEnquadramento.iTipoIPI, _
                objIPICodEnquadramento.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 216106

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216107

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objIPICodEnquadramento.sGrupoCST = tIPICodEnquadramento.sGrupoCST
    objIPICodEnquadramento.sDescCompleta = tIPICodEnquadramento.sDescCompleta
    objIPICodEnquadramento.iTipoIPI = tIPICodEnquadramento.iTipoIPI

    'Fecha Comando
    Call Comando_Fechar(lComando)

    IPICodEnquadramento_Le = SUCESSO

    Exit Function

Erro_IPICodEnquadramento_Le:

    IPICodEnquadramento_Le = gErr

    Select Case gErr

        Case 216105
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 216106, 216107
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPICODENQUADRAMENTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216108)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TribBaseDupla_Verifica(ByVal sTipo As String, ByVal sUFDestino As String, ByVal dtDataRef As Date, iBaseDupla As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim dtDataAux As Date

On Error GoTo Erro_TribBaseDupla_Verifica

    iBaseDupla = DESMARCADO

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 216105


    'Le a tabela TribBaseDupla
    lErro = Comando_Executar(lComando, "SELECT DataIni FROM TribBaseDupla WHERE Tipo= ? AND UF = ? AND DataIni <= ? ", _
                dtDataAux, sTipo, sUFDestino, dtDataRef)
    If lErro <> AD_SQL_SUCESSO Then gError 216106

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216107

    If lErro = AD_SQL_SUCESSO Then iBaseDupla = MARCADO

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TribBaseDupla_Verifica = SUCESSO

    Exit Function

Erro_TribBaseDupla_Verifica:

    TribBaseDupla_Verifica = gErr

    Select Case gErr

        Case 216105
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 216106, 216107
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBBASEDUPLA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216108)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
