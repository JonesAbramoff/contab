VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Tipos de Frete
Private Const TIPO_CIF = 0
Private Const TIPO_FOB = 1

Type typeImportCli
    lCodCliente As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodPadraoCobranca As Integer
    iCodMensagem As Integer
    iTabelaPreco As Integer
    lNumPagamentos As Long
    iProxCodFilial As Integer
    iCodFilial As Integer
    sFilialNome As String
    sFilialCGC As String
    sFilialInscEstadual As String
    sFilialInscMunicipal As String
    iFilialCodTransportadora As Integer
    sFilialObservacao1 As String
    sFilialContaContabil As String
    iFilialVendedor As Integer
    dFilialComissaoVendas As Double
    iFilialRegiao As Integer
    iFilialFreqVisitas As Integer
    dtFilialDataUltVisita As Date
    iFilialCodCobrador As Integer
    iFilialICMSBaseCalculoIPI As Integer
    lFilialRevendedor As Long
    sFilialTipoFrete As String
    sEndereco As String
    sBairro As String
    sCidade As String
    sSiglaEstado As String
    iCodigoPais As Integer
    sCEP As String
    sTelefone1 As String
    sTelefone2 As String
    sEmail As String
    sFax As String
    sContato As String
    sEndereco1 As String
    sBairro1 As String
    sCidade1 As String
    sSiglaEstado1 As String
    iCodigoPais1 As Integer
    sCEP1 As String
    sTelefone11 As String
    sTelefone21 As String
    sEmail1 As String
    sFax1 As String
    sContato1 As String
    sEndereco2 As String
    sBairro2 As String
    sCidade2 As String
    sSiglaEstado2 As String
    iCodigoPais2 As Integer
    sCEP2 As String
    sTelefone12 As String
    sTelefone22 As String
    sEmail2 As String
    sFax2 As String
    sContato2 As String
End Type

Public Function Cliente_Grava_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lCliente As Integer
Dim objClientesAF As ClassClientesAF

On Error GoTo Erro_Cliente_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192952
    Next

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClientesAF = objCliente.objInfoUsu
        
        objClientesAF.lCliente = objCliente.lCodigo
        objClientesAF.iTipoAssociado = objCliente.iTipo
        
        'Le a tabela ClientesAF
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM ClientesAF WHERE Empresa1 = ? AND MatriculaPrinc = ? AND TipoAssociado = ? AND Cliente <> ? ", _
                    0, lCliente, objClientesAF.iEmpresa1, objClientesAF.lMatriculaPrinc, objClientesAF.iTipoAssociado, objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 192949
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192950
        
        If lErro = AD_SQL_SUCESSO Then gError 192951

        'Le a tabela ClientesAF
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM ClientesAF WHERE Cliente= ? ", _
                    0, lCliente, objClientesAF.lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 192953
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192954
    
        'Se existir => UPDATE, senão => INSERT
        If lErro = AD_SQL_SUCESSO Then
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 192955
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ClientesAF SET TipoAssociado= ? , MatriculaPrinc= ? , StatusAssociado= ? , MatriculaSec= ? , " & _
                    "MatriculaAF= ? , MatriculaFRG= ? , Empresa1= ? , Empresa2= ? , Sexo= ? , " & _
                    "RGOrgaoEmissor= ? , DataExpedicaoRG= ? , DataNascimento= ? , DataInscricao= ? , LocalTrabalho= ? , " & _
                    "DataAdmissaoFurnas= ? , Cargo= ? , DataAposINSS= ? , DataAposFRG= ? , TipoApos= ? , " & _
                    "DataConBenf= ? , NoBenef= ? , DataFalecimento= ? , ContrSoc= ? , Observacao1= ? , " & _
                    "Observacao2= ? , Benemerito= ? , Fundador= ? , DataUltAtualizacao= ?, MatriculaPrincDV = ? ", alComando(0), _
                    objClientesAF.iTipoAssociado, objClientesAF.lMatriculaPrinc, objClientesAF.iStatusAssociado, objClientesAF.lMatriculaSec, _
                    objClientesAF.lMatriculaAF, objClientesAF.lMatriculaFRG, objClientesAF.iEmpresa1, objClientesAF.iEmpresa2, objClientesAF.sSexo, _
                    objClientesAF.sRGOrgaoEmissor, objClientesAF.dtDataExpedicaoRG, objClientesAF.dtDataNascimento, objClientesAF.dtDataInscricao, objClientesAF.sLocalTrabalho, _
                    objClientesAF.dtDataAdmissaoFurnas, objClientesAF.sCargo, objClientesAF.dtDataAposINSS, objClientesAF.dtDataAposFRG, objClientesAF.iTipoApos, _
                    objClientesAF.dtDataConBenf, objClientesAF.sNoBenef, objClientesAF.dtDataFalecimento, objClientesAF.sContrSoc, objClientesAF.sObservacao1, _
                    objClientesAF.sObservacao2, objClientesAF.iBenemerito, objClientesAF.iFundador, objClientesAF.dtDataUltAtualizacao, objClientesAF.iMatriculaPrincDV)
            If lErro <> AD_SQL_SUCESSO Then gError 192956
    
        Else
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO ClientesAF( Cliente, TipoAssociado, MatriculaPrinc, StatusAssociado, " & _
                    "MatriculaSec, MatriculaAF, MatriculaFRG, Empresa1, Empresa2, " & _
                    "Sexo, RGOrgaoEmissor, DataExpedicaoRG, DataNascimento, DataInscricao, " & _
                    "LocalTrabalho, DataAdmissaoFurnas, Cargo, DataAposINSS, DataAposFRG, " & _
                    "TipoApos, DataConBenf, NoBenef, DataFalecimento, ContrSoc, " & _
                    "Observacao1, Observacao2, Benemerito, Fundador, DataUltAtualizacao, MatriculaPrincDV) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objClientesAF.lCliente, objClientesAF.iTipoAssociado, objClientesAF.lMatriculaPrinc, objClientesAF.iStatusAssociado, _
                    objClientesAF.lMatriculaSec, objClientesAF.lMatriculaAF, objClientesAF.lMatriculaFRG, objClientesAF.iEmpresa1, objClientesAF.iEmpresa2, _
                    objClientesAF.sSexo, objClientesAF.sRGOrgaoEmissor, objClientesAF.dtDataExpedicaoRG, objClientesAF.dtDataNascimento, objClientesAF.dtDataInscricao, _
                    objClientesAF.sLocalTrabalho, objClientesAF.dtDataAdmissaoFurnas, objClientesAF.sCargo, objClientesAF.dtDataAposINSS, objClientesAF.dtDataAposFRG, _
                    objClientesAF.iTipoApos, objClientesAF.dtDataConBenf, objClientesAF.sNoBenef, objClientesAF.dtDataFalecimento, objClientesAF.sContrSoc, _
                    objClientesAF.sObservacao1, objClientesAF.sObservacao2, objClientesAF.iBenemerito, objClientesAF.iFundador, objClientesAF.dtDataUltAtualizacao, objClientesAF.iMatriculaPrincDV)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192957
    
        End If
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Grava_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Grava_Customizado:

    Cliente_Grava_Customizado = gErr

    Select Case gErr

        Case 192949, 192950, 192953, 192954
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)

        Case 192951
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTESAF_REPETIDO", gErr, objClientesAF.iEmpresa1, objClientesAF.lMatriculaPrinc, objCliente.iTipo)

        Case 192952
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192955
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CLIENTESAF", gErr)

        Case 192956
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CLIENTESAF", gErr)

        Case 192957
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLIENTESAF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192940)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Cliente_Exclui_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Integer

On Error GoTo Erro_Cliente_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192943
    Next

    'Le a tabelaClientesAF
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM ClientesAF WHERE Cliente= ? ", _
                0, lCliente, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192944

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192945

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192946
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ClientesAF", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 192947
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Exclui_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Exclui_Customizado:

    Cliente_Exclui_Customizado = gErr

    Select Case gErr

        Case 192943
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192944, 192945
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)
        
        Case 192946
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CLIENTESAF", gErr)

        Case 192947
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTESAF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192948)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_AFEmpresas() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_AFEmpresas
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194011
    
    'Busca nos clientes empresas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Cliente, Empresa1 FROM ClientesAF WHERE ((Empresa1 <> 0 AND Empresa1 NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?)) OR (Empresa2 <> 0 AND Empresa2 NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?))) ", lCodigo, iTipo, CAMPOSGENERICOS_AF_EMPRESAS, CAMPOSGENERICOS_AF_EMPRESAS)
    If lErro <> AD_SQL_SUCESSO Then gError 194012
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194013
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 194014
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_AFEmpresas = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_AFEmpresas:

    CamposGenericos_ValidaExclusao_AFEmpresas = gErr
    
    Select Case gErr
    
        Case 194011
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 194012, 194013
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_AF_EMPRESAS)
        
        Case 194014
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_AF_EMPRESA", gErr, iTipo, lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194015)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function CamposGenericos_ValidaExclusao_TipoApos() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TipoApos
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194016
    
    'Busca nos clientes empresas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Cliente, TipoApos FROM ClientesAF WHERE TipoApos <> 0 AND TipoApos NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iTipo, CAMPOSGENERICOS_AF_TIPOAPOS)
    If lErro <> AD_SQL_SUCESSO Then gError 194017
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194018
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 194019
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_TipoApos = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TipoApos:

    CamposGenericos_ValidaExclusao_TipoApos = gErr
    
    Select Case gErr
    
        Case 194016
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 194017, 194018
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_AF_EMPRESAS)
        
        Case 194019
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_AF_TIPOAPOS", gErr, iTipo, lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194020)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function Importa_CR_AF(ByVal iFilialEmpresa As Integer, ByVal iCobrador As Integer, ByVal sNomeDiretorio As String, ByVal sNomeArquivo As String, ByVal dtDataManual As Date) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim colRetornoArq As New Collection
Dim colErros As New Collection
Dim objRetorno As ClassDescAposPenAF
Dim lTransacao As Long
Dim lNumTitRecAuto As Long
Dim objTitRec As ClassTituloReceber
Dim colTitRec As New Collection
Dim dtData As Date
Dim colBxParcRec As New ColBaixaParcRec
Dim bImportado As Boolean
Dim dtDataImp As Date
Dim sUsuarioImp As String
Dim dValor As Double
Dim vbMsgRst As VbMsgBoxResult
Dim lNumIntBaixa As Long
Dim colParcRec As colParcelaReceber

On Error GoTo Erro_Importa_CR_AF

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194051
    Next
    
    'Verifica se o arquivo ainda não foi importado
    lErro = Importa_CR_Le(alComando(1), sNomeArquivo, bImportado, dtDataImp, sUsuarioImp)
    If lErro <> SUCESSO Then gError 194059
    
    If bImportado Then gError 194060 'Arquivo já importado anteriormente (dtDataImp, sUsuarioImp)
    
    'Le os registros do arquivo
    lErro = Importa_CR_Le_Arquivo(sNomeDiretorio, sNomeArquivo, colRetornoArq, dtData, dValor, iCobrador)
    If lErro <> SUCESSO Then gError 194053
    
    If dtDataManual <> DATA_NULA Then
        If dtDataManual <> dtData Then
            vbMsgRst = Rotina_Aviso(vbYesNo, "AVISO_DATA_INF_DIF_DATA_CALC", dtDataManual, dtData)
            If vbMsgRst = vbNo Then gError 194053
            dtData = dtDataManual
        End If
    End If
    
    vbMsgRst = Rotina_Aviso(vbYesNo, "AVISO_CONFIMACAO_MOVCC_INCLUSAO_AF", dtData, Format(dValor, "STANDARD"), iCobrador)
    If vbMsgRst = vbNo Then gError 194062
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194052
       
    For Each objRetorno In colRetornoArq
        
        'Só testa para ver se todos os clientes atendem a condição antes de gerar o documento de fato
        lErro = Gera_CR(alComando(2), iFilialEmpresa, EMPRESA_FURNAS, objRetorno.lMatricula, lNumTitRecAuto, dtData, objRetorno.dValor, objTitRec, colErros, True)
        If lErro <> SUCESSO Then gError 194054
       
    Next
       
    For Each objRetorno In colRetornoArq
    
        Set objTitRec = New ClassTituloReceber
        Set colParcRec = New colParcelaReceber
    
        'Obtém a numeração
        lErro = CF("Config_ObterAutomatico_EmTrans", "FATAFConfig", "NUM_PROX_TITULO_REC", "TitulosRecTodos", "NumTitulo", lNumTitRecAuto)
        If lErro <> SUCESSO Then gError 194095
        
        'Gera, Grava e Contabiliza os títulos a receber
        lErro = Gera_CR(alComando(2), iFilialEmpresa, EMPRESA_FURNAS, objRetorno.lMatricula, lNumTitRecAuto, dtData, objRetorno.dValor, objTitRec, colErros)
        'If lErro <> SUCESSO And lErro <> 194043 And lErro <> 194045 And lErro <> 194046 And lErro <> 194047 Then gError 194054
        If lErro <> SUCESSO Then gError 194054
        
        lErro = CF("TituloReceber_Le", objTitRec)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 194098
        
        'Lê as Parcelas a Receber vinculadas ao Título
        lErro = CF("ParcelasReceber_Le", objTitRec, colParcRec)
        If lErro <> SUCESSO Then gError 194099
        
        colTitRec.Add objTitRec
        Call colBxParcRec.Add(0, 0, colParcRec.Item(1).lNumIntDoc, 0, STATUS_LANCADO, 0, 0, 0, objTitRec.dValor, objTitRec.dValor, iCobrador)
    
    Next
    
    'Gera a baixa
    lErro = Baixa_CR(colBxParcRec, colTitRec, iCobrador, dtData, dValor, lNumIntBaixa, colErros)
    If lErro <> SUCESSO Then gError 194055
    
    'Grava o arquivo para indicar que ele já foi importado
    lErro = Importa_CR_Grava(alComando(3), sNomeArquivo, lNumIntBaixa, iCobrador, dtData)
    If lErro <> SUCESSO Then gError 194061
    
    If colErros.Count <> 0 Then
    
        gError 194056
    End If

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194057
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Importa_CR_AF = SUCESSO

    Exit Function

Erro_Importa_CR_AF:

    Importa_CR_AF = gErr

    Select Case gErr

        Case 194051
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194052
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 194053 To 194055, 194059, 194061, 194062, 194095, 194098, 194099
        
        Case 194056 'Deve ser tratado na rotina chamadora -> Chamar relatório de erros
        
        Case 194057
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 194060 'Arquivo já importado
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_LIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194058)

    End Select

    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Gera_CR(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iEmpresa As Integer, ByVal lMatricula As Long, ByVal lNumTitulo As Long, ByVal dtDataVenc As Date, ByVal dValor As Double, ByVal objTitRec As ClassTituloReceber, ByVal colErros As Collection, Optional bSoTestaCliente As Boolean = False) As Long

Dim lErro As Long
Dim lCliente As Long
Dim lClienteAux As Long
Dim objCliente As New ClassCliente
Dim objClienteAF As New ClassClientesAF
Dim objParcelaReceber As New ClassParcelaReceber
Dim colParcelaReceber As New colParcelaReceber
Dim colComissaoEmissao As New colComissao
Dim colcolComissao As New colcolComissao
Dim colComissao As New colComissao
Dim colcolDesconto As New colcolDesconto
Dim colDesconto As New colDesconto
Dim objContabil As New ClassContabil
Dim objCtb As ClassCRCtbAF
Dim objContabAutomatica As ClassContabAutomatica
Dim lNumAuto As Long

On Error GoTo Erro_Gera_CR

    'Le a tabela ClientesAF
    lErro = Comando_Executar(lComando, "SELECT C2.Cliente FROM Clientes AS C1, ClientesAF AS C2 WHERE C2.Cliente = C1.Codigo AND C2.Empresa1 = ? AND C2.MatriculaPrinc = ? AND C1.Ativo = ? ", _
                lClienteAux, iEmpresa, lMatricula, CLIENTE_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 194041

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194042
    
    'Não Achou o cliente em questão
    If lErro <> AD_SQL_SUCESSO Then gError 194043
    
    lCliente = lClienteAux
    
    'Busca Proximo
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194044
    
    'Tem mais de um cliente ativo para essa empresa\matricula
    If lErro = AD_SQL_SUCESSO Then gError 194045
    
    If Not bSoTestaCliente Then

        objTitRec.lCliente = lCliente
        objTitRec.iFilial = FILIAL_MATRIZ
        objTitRec.lNumTitulo = lNumTitulo
        objTitRec.dtDataEmissao = dtDataVenc
        objTitRec.dValor = dValor
        objTitRec.dSaldo = dValor
        objTitRec.iFilialEmpresa = iFilialEmpresa
        objTitRec.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER
        objTitRec.iNumParcelas = 1
    
        objParcelaReceber.iNumParcela = 1
        objParcelaReceber.dtDataVencimento = dtDataVenc
        objParcelaReceber.dtDataVencimentoReal = dtDataVenc
        objParcelaReceber.dValor = dValor 'dvalbru
        objParcelaReceber.dValorOriginal = dValor 'dvalbru
        objParcelaReceber.dtDataCredito = DATA_NULA
        objParcelaReceber.dtDataDepositoCheque = DATA_NULA
        objParcelaReceber.dtDataEmissaoCheque = DATA_NULA
        objParcelaReceber.dtDataPrevReceb = DATA_NULA
        objParcelaReceber.dtDataProxCobr = DATA_NULA
        objParcelaReceber.dtDataTransacaoCartao = DATA_NULA
        objParcelaReceber.dtDesconto1Ate = DATA_NULA
        objParcelaReceber.dtDesconto2Ate = DATA_NULA
        objParcelaReceber.dtDesconto3Ate = DATA_NULA
        objParcelaReceber.dtValidadeCartao = DATA_NULA
        
        colParcelaReceber.AddObj objParcelaReceber
    
        colcolDesconto.Add colDesconto
        colcolComissao.Add colComissao
        
        'Grava o Título a receber
        lErro = CF("TituloReceber_Grava_EmTrans", objTitRec, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, Nothing)
        If lErro <> SUCESSO Then gError 194046
        
        Set objContabAutomatica = New ClassContabAutomatica
        Set objCtb = New ClassCRCtbAF
        
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
        If lErro <> SUCESSO Then gError 194047
        
        lErro = objContabAutomatica.GeraContabilizacao(objTitRec)
        If lErro <> SUCESSO Then gError 194048
        
        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then gError 194049
    
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
    End If

    Gera_CR = SUCESSO

    Exit Function

Erro_Gera_CR:

    Gera_CR = gErr

    Select Case gErr
  
        Case 194041, 194042, 194044
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)
        
        Case 194043 'Cliente ativo não encontrado
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_CLIENTE_ATIVO_AF_PARA_MATRICULA", gErr, iEmpresa, lMatricula)
            colErros.Add "Não existe cliente ativo para Empresa " & CStr(iEmpresa) & " com a Matrícula " & CStr(lMatricula) & "."

        Case 194045 'Mais de um Cliente ativo
            Call Rotina_Erro(vbOKOnly, "ERRO_MAIS_DE_UM_CLIENTE_ATIVO_AF_PARA_MATRICULA", gErr, iEmpresa, lMatricula)
            colErros.Add "Existe mais de um cliente ativo para Empresa " & CStr(iEmpresa) & " com a Matrícula " & CStr(lMatricula) & "."
        
        Case 194046 'Erro na gravação
            'Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_TITULOREC_AF", gErr, iEmpresa, lMatricula)
            colErros.Add "Não foi possível gravar o Título para o Cliente da Empresa " & CStr(iEmpresa) & " e Matrícula " & CStr(lMatricula) & "."
        
        Case 194047 To 194049 'Erro na contabilização
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONTABILIZACAO_TITULOREC_AF", gErr, iEmpresa, lMatricula)
            colErros.Add "Não foi possível contabilizar o Título para o Cliente da Empresa " & CStr(iEmpresa) & " e Matrícula " & CStr(lMatricula) & "."

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194050)

    End Select

    Exit Function

End Function

Private Function Baixa_CR(ByVal colBaixaParcReceber As ColBaixaParcRec, ByVal colTitRec As Collection, ByVal iCobrador As Integer, ByVal dtData As Date, ByVal dValor As Double, ByVal lNumIntBaixa As Long, ByVal colErros As Collection) As Long

Dim lErro As Long
Dim objCtb As New ClassBxCRCtbAF
Dim objContabAutomatica As New ClassContabAutomatica
Dim objCobrador As New ClassCobrador
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim objMovCCI As New ClassMovContaCorrente

On Error GoTo Erro_Baixa_CR
    
    objCobrador.iCodigo = iCobrador
    
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO Then gError 194063
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 194064
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then gError 194065
        
    With objMovCCI
    
        .iFilialEmpresa = objContaCorrente.iFilialEmpresa
        .iCodConta = objContaCorrente.iCodigo
        .iExcluido = NAO_EXCLUIDO
        .iTipoMeioPagto = RECEBIMENTO_EM_DINHEIRO
        .dtDataBaixa = dtData
        .dtDataMovimento = dtData
        .dtDataContabil = dtData
        .iConciliado = NAO_CONCILIADO
        .dValor = dValor
        .iTipo = MOVCCI_RECEBIMENTO_TITULO
        
    End With
    
    'Chama BaixaReceber_Grava_Recebimento dentro de transação
    lErro = CF("BaixaRec_Grava_Recebimento_Trans", colBaixaParcReceber, objMovCCI, Nothing, 0, 0)
    If lErro <> SUCESSO Then gError 194066
    
    lNumIntBaixa = colBaixaParcReceber.Item(1).lNumIntBaixa
    
    'gravar a contabilizacao
    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", dtData, dtData)
    If lErro <> SUCESSO Then gError 194067
    
    lErro = objContabAutomatica.GeraContabilizacao(colBaixaParcReceber, colTitRec, lNumIntBaixa, dtData, objMovCCI)
    If lErro <> SUCESSO Then gError 194068
    
    lErro = objContabAutomatica.Finaliza_Contab()
    If lErro <> SUCESSO Then gError 194069

    Set objContabAutomatica = Nothing
    Call objCtb.Encerrando
    Set objCtb = Nothing

    Baixa_CR = SUCESSO

    Exit Function

Erro_Baixa_CR:

    Baixa_CR = gErr

    Select Case gErr
    
        Case 194063, 194064
    
        Case 194065
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, objContaCorrente.iCodigo)
            colErros.Add "Não foi possível ler a conta corrente código " & CStr(objContaCorrente.iCodigo) & "."
        
        Case 194066 'Erro na gravação
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONTABILIZACAO_TITULOREC_AF", gErr, iEmpresa, lMatricula)
            colErros.Add "Não foi possível gravar a baixa dos títulos."
       
        Case 194067 To 194069 'Erro na contabilização
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONTABILIZACAO_BAIXA_TITULOREC_AF", gErr)
            colErros.Add "Não foi possível contabilizar a baixa dos títulos."

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194070)

    End Select

    Exit Function

End Function

Private Function Importa_CR_Le(ByVal lComando As Long, ByVal sNomeArquivo As String, bImportado As Boolean, dtDataImp As Date, sUsuarioImp As String) As Long

Dim lErro As Long

On Error GoTo Erro_Importa_CR_Le

    sUsuarioImp = String(STRING_USUARIO_NOMEREDUZIDO, 0)

    'Le a tabela ClientesAF
    lErro = Comando_Executar(lComando, "SELECT DataImportacao, Usuario FROM ArqImpCRAF WHERE NomeArquivo = ? ", dtDataImp, sUsuarioImp, sNomeArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 194081

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194082
    
    'Não Achou o arquivo em questão
    If lErro <> AD_SQL_SUCESSO Then
        bImportado = False
    Else
        bImportado = True
    End If

    Importa_CR_Le = SUCESSO

    Exit Function

Erro_Importa_CR_Le:

    Importa_CR_Le = gErr

    Select Case gErr
    
        Case 194081, 194082
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQIMPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194083)

    End Select

    Exit Function

End Function

Private Function Importa_CR_Grava(ByVal lComando As Long, ByVal sNomeArquivo As String, ByVal lNumIntBaixa As Long, ByVal iCobrador As Integer, ByVal dtDataRef As Date) As Long

Dim lErro As Long

On Error GoTo Erro_Importa_CR_Grava

    'Le a tabela ClientesAF
    lErro = Comando_Executar(lComando, "INSERT INTO ArqImpCRAF (NomeArquivo, DataImportacao, HoraImportacao, Usuario, NumIntDocBaixa, Cobrador,DataRef) VALUES (?,?,?,?,?,?,?) ", sNomeArquivo, Date, CDbl(Time), gsUsuario, lNumIntBaixa, iCobrador, dtDataRef)
    If lErro <> AD_SQL_SUCESSO Then gError 194084

    Importa_CR_Grava = SUCESSO

    Exit Function

Erro_Importa_CR_Grava:

    Importa_CR_Grava = gErr

    Select Case gErr
    
        Case 194084
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ARQIMPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194085)

    End Select

    Exit Function

End Function

Private Function Importa_CR_Le_Arquivo(ByVal sNomeDiretorio As String, ByVal sNomeArquivo As String, ByVal colRetornoArq As Collection, dtData As Date, dValor As Double, ByVal iCobrador As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double
Dim objFSO As New FileSystemObject
Dim objTS As TextStream
Dim bArqAberto As Boolean
Dim sRegistro As String
Dim objDescAposPen As ClassDescAposPenAF
Dim iPos As Integer
Dim dtDataAux As Date
Dim iNumDiasUteis As Integer
Dim dtDataAuxReal As Date
Dim iLinha As Integer
Dim sDataRef As String

On Error GoTo Erro_Importa_CR_Le_Arquivo
       
    bArqAberto = False
    
    'abrir arquivo texto
    Set objTS = objFSO.OpenTextFile(sNomeDiretorio & sNomeArquivo, ForReading, False)
    bArqAberto = True

    dValorAux = 0
    iLinha = 0
    
    'Até chegar ao fim do arquivo
    Do While Not objTS.AtEndOfLine
    
        'Busca o próximo registro do arquivo
        sRegistro = objTS.ReadLine
        iLinha = iLinha + 1
         
        Set objDescAposPen = New ClassDescAposPenAF
        
        If iCobrador = EMPRESA_FRG Then
         
            iPos = 1
             
            objDescAposPen.lMatricula = StrParaLong(Mid(sRegistro, iPos, 6))
            iPos = iPos + 6
            
            objDescAposPen.iDV = StrParaInt(Mid(sRegistro, iPos, 1))
            iPos = iPos + 1
             
            objDescAposPen.iAno = StrParaInt(Mid(sRegistro, iPos, 4))
            iPos = iPos + 4
            
            objDescAposPen.iMes = StrParaInt(Mid(sRegistro, iPos, 2))
            iPos = iPos + 2
            
            'objDescAposPen.dValor = StrParaDbl(Mid(sRegistro, iPos, 13)) / 100
            'iPos = iPos + 13
            
            iPos = iPos + 3 'código utilizado pela Fundação que não tem efeito para Após-Furnas
            
            objDescAposPen.dValor = StrParaDbl(Mid(sRegistro, iPos, 8)) / 100
            iPos = iPos + 8
            
            colRetornoArq.Add objDescAposPen
            
        ElseIf iCobrador = EMPRESA_FURNAS Then
        
            If iLinha = 1 Then
                sDataRef = left(sRegistro, 6)
            Else
                     
                iPos = 1
                 
                objDescAposPen.lMatricula = StrParaLong(Mid(sRegistro, iPos, 6))
                iPos = iPos + 6
                
                objDescAposPen.iDV = StrParaInt(Mid(sRegistro, iPos, 1))
                iPos = iPos + 1
                 
                objDescAposPen.iMes = StrParaInt(right(sDataRef, 2))
                objDescAposPen.iAno = StrParaInt(left(sDataRef, 4))
                
                iPos = iPos + 19
                
                objDescAposPen.dValor = StrParaDbl(Mid(sRegistro, iPos, 11)) / 100
                iPos = iPos + 12
                
                If InStr(1, sRegistro, "FORA FOLHA") = 0 Or objDescAposPen.lMatricula = 14989 Then
                    colRetornoArq.Add objDescAposPen
                End If
                
            End If
            
        End If
         
        dValorAux = dValorAux + objDescAposPen.dValor
         
         
    Loop
    
    dValor = dValorAux
    
    'fechar arquivo texto
    objTS.Close
    bArqAberto = False
    
    If colRetornoArq.Count > 0 Then
    
        Set objDescAposPen = colRetornoArq.Item(1)
    
        dtDataAux = StrParaDate("01/" & CStr(objDescAposPen.iMes) & "/" & CStr(objDescAposPen.iAno))
        
        'Primeiro dia do mês seguinte
        dtDataAux = DateAdd("m", 1, dtDataAux)
        
        iNumDiasUteis = 0
        Do While iNumDiasUteis <> 2
        
            dtDataAux = DateAdd("d", -1, dtDataAux)
        
            lErro = CF("DataVencto_Real", dtDataAux, dtDataAuxReal)
            If lErro <> SUCESSO Then gError 194097
            
            If dtDataAux = dtDataAuxReal Then
                iNumDiasUteis = iNumDiasUteis + 1
            End If
        
        Loop
    
    Else
        dtDataAux = DATA_NULA
    End If
    
    dtData = dtDataAux

    Importa_CR_Le_Arquivo = SUCESSO

    Exit Function

Erro_Importa_CR_Le_Arquivo:

    Importa_CR_Le_Arquivo = gErr

    Select Case gErr
    
        Case 194097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194086)

    End Select
    
    'fechar aquivo texto
    If bArqAberto Then objTS.Close

    Exit Function

End Function

Public Function Exporta_Desc_Apos(ByVal dtData As Date, ByVal sNomeArquivo As String, ByVal dValorContribSoc As Double) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim sRegistro As String
Dim objFSO As New FileSystemObject
Dim objTS As TextStream
Dim bCriouArq As Boolean
Dim dValorTotal As Double
Dim lMatriculaPrinc As Long
Dim iMatriculaPrincDV As Long
Dim iAno As Integer
Dim iAnoAux As Integer
Dim iMes As Integer
Dim lCliente As Long
Dim iTipoCli As Integer
Dim iFoiArq As Integer
Dim sContrSoc As String

On Error GoTo Erro_Exporta_Desc_Apos

    bCriouArq = False

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194100
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194132
    
    Set objTS = objFSO.OpenTextFile(sNomeArquivo, ForWriting, True)
    bCriouArq = True
    
    iAno = Year(dtData)
    iMes = Month(dtData)
    
    'Le os clientes que já foram exportados para esse mês\ano
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Ano FROM ArqExpCRAFDet WHERE Ano = ? AND Mes = ?", 0, iAnoAux, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 194134

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194135
    
    Do While lErro = AD_SQL_SUCESSO
                
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> SUCESSO Then gError 194136

        'Exclui
        lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ArqExpCRAFDet ", alComando(3))
        If lErro <> AD_SQL_SUCESSO Then gError 194137
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194138
            
    Loop
    
    sContrSoc = String(STRING_CLIENTESAF_CONTRSOC, 0)
    
    'LE OS ASSOCIADOS QUE TEM QUE TER DESCONTO
    lErro = Comando_Executar(alComando(0), "SELECT CAF.MatriculaPrinc, CAF.MatriculaPrincDV, C.Codigo, C.Tipo, CAF.ContrSoc FROM Clientes AS C, ClientesAF AS CAF WHERE C.Codigo = CAF.Cliente AND C.Ativo = ? AND CAF.Empresa1 = ? ORDER BY CAF.MatriculaPrinc", lMatriculaPrinc, iMatriculaPrincDV, lCliente, iTipoCli, sContrSoc, MARCADO, EMPRESA_FURNAS)
    If lErro <> AD_SQL_SUCESSO Then gError 194101

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194102

    objTS.WriteLine "0" & Format(dtData, "YYYYMM")
    
    dValorTotal = 0
    Do While lErro = AD_SQL_SUCESSO
    
        sRegistro = "1574" & FormataCpoNum(lMatriculaPrinc, 6) & FormataCpoNum(iMatriculaPrincDV, 1) & FormataCpoValor(dValorContribSoc, 11) & SEPARADOR
        
        iFoiArq = DESMARCADO
        If iTipoCli = 3 And sContrSoc = "Débito Automático Fl. Pgto." Then
            objTS.WriteLine sRegistro
            iFoiArq = MARCADO
            dValorTotal = dValorTotal + dValorContribSoc
        End If

        'Insere no detalhe do arquivo
        lErro = Comando_Executar(alComando(5), "INSERT INTO ArqExpCRAFDet (Ano,Mes,Cliente, FoiArq) VALUES (?,?,?,?)", iAno, iMes, lCliente, iFoiArq)
        If lErro <> AD_SQL_SUCESSO Then gError 194139
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194103
        
    Loop
    
    objTS.WriteLine "99999999999" & FormataCpoValor(dValorTotal, 11) & SEPARADOR
    
    objTS.Close
       
    'INSERE DATA E O VALOR DA MENSALIDADE EM UMA TABELA (Usuário, DataGer, HoraGer, NomeArq)
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Ano FROM ArqExpCRAF WHERE Ano = ? AND Mes = ?", 0, iAnoAux, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 194127

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194128
    
    If lErro <> AD_SQL_SUCESSO Then
        
        'Insere arquivo
        lErro = Comando_Executar(alComando(2), "INSERT INTO ArqExpCRAF (NomeArquivo,DataGeracao,HoraGeracao,Usuario,Ano,Mes,ValorContribSoc) VALUES (?,?,?,?,?,?,?)", sNomeArquivo, Date, CDbl(Time), gsUsuario, iAno, iMes, dValorContribSoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194129
        
    Else
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 194130

        'Altera o arquivo
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ArqExpCRAF SET NomeArquivo= ? , DataGeracao= ? , HoraGeracao= ? , Usuario= ? , ValorContribSoc = ? ", alComando(1), sNomeArquivo, Date, CDbl(Time), gsUsuario, dValorContribSoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194131
        
    End If
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194133
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exporta_Desc_Apos = SUCESSO

    Exit Function

Erro_Exporta_Desc_Apos:

    Exporta_Desc_Apos = gErr

    Select Case gErr

        Case 194100
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 194101, 194102, 194103
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)
        
        Case 194127, 194128, 194134, 194135, 194138
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQEXPCRAF", gErr)
        
        Case 194129, 194139
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ARQEXPCRAF", gErr)
        
        Case 194130, 194136
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ARQEXPCRAF", gErr)
        
        Case 194131, 194137
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ARQEXPCRAF", gErr)

        Case 194132
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 194133
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194104)

    End Select
    
    Call Transacao_Rollback
    
    If bCriouArq Then objFSO.DeleteFile (sNomeArquivo)

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelInadAF_Prepara(lNumIntRel As Long, ByVal lClienteDe As Long, ByVal lClienteAte As Long, ByVal iAnoDe As Integer, ByVal iAnoAte As Integer, ByVal iMesDe As Integer, ByVal iMesAte As Integer) As Long

Dim lErro As Long
Dim lErro2 As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lCliente As Long
Dim lClienteAnt As Long
Dim iAno As Integer
Dim iMes As Integer
Dim dValorContribSoc As Double
Dim dValorContribSocTotal As Double
Dim dValorTitRec As Double
Dim sMes As String
Dim iCount As Integer
Dim lMatricula As Long
Dim sEmpresa As String
Dim sCgc As String
Dim sNomeCliente As String
Dim lMatriculaAnt As Long
Dim sEmpresaAnt As String
Dim sCGCAnt As String
Dim sNomeClienteAnt As String
Dim iTipoCli As Integer
Dim iTipoCliAnt As Integer
Dim dtDataAuxIni As Date, dtDataAuxFim As Date

On Error GoTo Erro_RelInadAF_Prepara

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194140
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194141
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATAFConfig", "NUM_PROX_REL_INADAF", lNumIntRel)
    If lErro <> SUCESSO Then gError 194142
    
    sEmpresa = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    sCgc = String(STRING_CGC, 0)
    sNomeCliente = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT D.Cliente, C.Ano, C.Mes, C.ValorContribSoc, CAF.ClienteNomeRed, CAF.MatriculaPrinc, CAF.Empresa1Nome, CAF.ClienteCGC, CAF.TipoClienteCod FROM ArqExpCRAFDet AS D, ArqExpCRAF AS C, ClientesFiliais_AF AS CAF WHERE D.Cliente = CAF.ClienteCod AND D.Ano = C.Ano AND D.Mes = C.Mes AND (D.Cliente >= ? OR ? = 0 ) AND (D.Cliente <= ? OR ? = 0 ) AND (C.Ano > ? OR (C.Ano = ? AND C.Mes >= ?) OR ? = 0) AND (C.Ano < ? OR (C.Ano = ? AND C.Mes <= ?) OR ? = 0) ORDER BY D.Cliente, C.Ano, C.Mes", lCliente, iAno, iMes, dValorContribSoc, sNomeCliente, lMatricula, sEmpresa, sCgc, iTipoCli, lClienteDe, lClienteDe, lClienteAte, lClienteAte, iAnoDe, iAnoDe, iMesDe, iAnoDe, iAnoAte, iAnoAte, iMesAte, iAnoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 194143
    
   'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194144
    
    lClienteAnt = lCliente
    lMatriculaAnt = lMatricula
    sEmpresaAnt = sEmpresa
    sCGCAnt = sCgc
    sNomeClienteAnt = sNomeCliente
    iTipoCliAnt = iTipoCli

    Do While lErro = AD_SQL_SUCESSO
    
        dtDataAuxIni = StrParaDate("12/" & Format(iMes, "00") & "/" & CStr(iAno))
        If iMes <> 12 Then
            dtDataAuxFim = StrParaDate("11/" & Format(iMes + 1, "00") & "/" & CStr(iAno))
        Else
            dtDataAuxFim = StrParaDate("11/01/" & CStr(iAno + 1))
        End If

        'lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) FROM TitulosRecTodos WHERE MONTH(DataEmissao) = ? AND YEAR(DataEmissao) = ? AND Cliente = ?", dValorTitRec, iMes, iAno, lCliente)
        lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) FROM TitulosRecTodos WHERE DataEmissao BETWEEN ? AND ? AND Cliente = ?", dValorTitRec, dtDataAuxIni, dtDataAuxFim, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 194145
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194146
        
        'Não encontrou o a receber para esse cliente nesse mês
        If dValorTitRec < DELTA_VALORMONETARIO Then
        
            Call MesNome(iMes, sMes)
        
            'Insere arquivo
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelInadAFDet (NumIntRel,Cliente,Ano,Mes,MesNome,ValorDevido) VALUES (?,?,?,?,?,?)", lNumIntRel, lCliente, iAno, iMes, sMes, dValorContribSoc)
            If lErro <> AD_SQL_SUCESSO Then gError 194147
        
            dValorContribSocTotal = dValorContribSocTotal + dValorContribSoc
            iCount = iCount + 1
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194148
        
        If lClienteAnt <> lCliente Or lErro = AD_SQL_SEM_DADOS Then
        
            If iCount > 0 Then
            
                'Insere arquivo
                lErro2 = Comando_Executar(alComando(3), "INSERT INTO RelInadAF (NumIntRel,Cliente,Matricula,Empresa,CGC,NomeCliente,MesesDevidos,TotalDevido, TipoCli) VALUES (?,?,?,?,?,?,?,?,?)", lNumIntRel, lClienteAnt, lMatriculaAnt, sEmpresaAnt, sCGCAnt, sNomeClienteAnt, iCount, dValorContribSocTotal, iTipoCliAnt)
                If lErro2 <> AD_SQL_SUCESSO Then gError 194149
                
            End If
            
            lClienteAnt = lCliente
            lMatriculaAnt = lMatricula
            sEmpresaAnt = sEmpresa
            sCGCAnt = sCgc
            sNomeClienteAnt = sNomeCliente
            iTipoCliAnt = iTipoCli
            dValorContribSocTotal = 0
            iCount = 0
    
        End If
        
    Loop

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194150
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelInadAF_Prepara = SUCESSO

    Exit Function

Erro_RelInadAF_Prepara:

    RelInadAF_Prepara = gErr

    Select Case gErr

        Case 194140
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194141
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 194142
        
        Case 194143, 194144, 194148
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQEXPCRAF", gErr)

        Case 194145, 194146
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr)
            
        Case 194147
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_RELINADAFDET", gErr)

        Case 194149
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_RELINADAF", gErr)

        Case 194150
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194058)

    End Select

    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Importa_Clientes() As Long

Dim lErro As Long
Dim lComando As Long, lComando2 As Long
Dim lTransacao As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim colEndereco As New Collection, objEndereco As ClassEndereco
Dim lCodigo As Long
Dim EMPRESA_1 As String, TIPO_ASSOC As String, MATR_PRINC As Double, STAT_ASSOC As String, EMPRESA_2 As String, MATR_SEC As Double, MATR_AF As Double
Dim MATR_FRG  As Double, Nome As String, ENDERECO As String, BAIRRO As String, CIDADE As String, UF As String, CEP As String, TELEFONES As String, EMAIL As String, SEXO As String, IDENTIDADE As String, ORG_EMISS As String
Dim DT_EXPED As String, CPF As String, DT_NASC As Date, DT_INSCR As Date, LOC_TRAB As String, DT_ADMIT_FURNAS As Date, CARGO As String, DT_APOS_INSS As String, DT_APOSFRG As Date
Dim TIPO_APOS As String, DT_CONBENF As String, NO_BENEF As String, DT_FALEC As Date, CONTR_SOC As String, OBS_1 As String, OBS_2 As String, BENEMERITO As String, FUNDADOR As String, DT_ULTATU As Date
Dim sSQL As String, objClientesAF As ClassClientesAF, iPosLixo As Integer
Dim dvmatprinc As Long, lMatriculaPrinc As Long, iFimNome As Integer
Dim sAux As String

On Error GoTo Erro_Importa_Clientes

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76329
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 76329
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76342
    
    sSQL = "SELECT EMPRESA_1,TIPO_ASSOC,MATR_PRINC,STAT_ASSOC,EMPRESA_2,MATR_SEC,MATR_AF," & _
        "MATR_FRG , Nome, ENDERECO, BAIRRO, CIDADE, UF, CEP, TELEFONES, EMAIL, SEXO, IDENTIDADE, ORG_EMISS, " & _
        "DT_EXPED , CPF, DT_NASC, DT_INSCR, LOC_TRAB, [DT_ADMIT-FURNAS], CARGO, DT_APOS_INSS, DT_APOSFRG, " & _
        "TIPO_APOS , DT_CONBENF, NO_BENEF, DT_FALEC, CONTR_SOC, OBS_1, OBS_2, BENEMERITO, FUNDADOR, DT_ULTATU" & _
        " FROM ASSOC ORDER BY MATR_AF, DT_NASC"
    
    EMPRESA_1 = String(512, 0)
    TIPO_ASSOC = String(512, 0)
    STAT_ASSOC = String(512, 0)
    EMPRESA_2 = String(512, 0)
    Nome = String(512, 0)
    ENDERECO = String(512, 0)
    BAIRRO = String(512, 0)
    CIDADE = String(512, 0)
    UF = String(512, 0)
    CEP = String(512, 0)
    TELEFONES = String(512, 0)
    EMAIL = String(512, 0)
    SEXO = String(512, 0)
    IDENTIDADE = String(512, 0)
    ORG_EMISS = String(512, 0)
    CPF = String(512, 0)
    LOC_TRAB = String(512, 0)
    CARGO = String(512, 0)
    TIPO_APOS = String(512, 0)
    NO_BENEF = String(512, 0)
    CONTR_SOC = String(512, 0)
    OBS_1 = String(512, 0)
    OBS_2 = String(512, 0)
    BENEMERITO = String(512, 0)
    FUNDADOR = String(512, 0)
        
    DT_EXPED = String(512, 0)
    DT_APOS_INSS = String(512, 0)
'    DT_FALEC = String(512, 0)
    DT_CONBENF = String(512, 0)
            
    lCodigo = 1000
    
    'Lê os registros da tabela de importacao
    lErro = Comando_Executar(lComando, sSQL, EMPRESA_1, TIPO_ASSOC, MATR_PRINC, STAT_ASSOC, EMPRESA_2, MATR_SEC, MATR_AF, _
        MATR_FRG, Nome, ENDERECO, BAIRRO, CIDADE, UF, CEP, TELEFONES, EMAIL, SEXO, IDENTIDADE, ORG_EMISS, _
        DT_EXPED, CPF, DT_NASC, DT_INSCR, LOC_TRAB, DT_ADMIT_FURNAS, CARGO, DT_APOS_INSS, DT_APOSFRG, _
        TIPO_APOS, DT_CONBENF, NO_BENEF, DT_FALEC, CONTR_SOC, OBS_1, OBS_2, BENEMERITO, FUNDADOR, DT_ULTATU)
    If lErro <> AD_SQL_SUCESSO Then gError 76330
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76331
    
    Do While lErro = AD_SQL_SUCESSO
            
        Set objCliente = New ClassCliente
        Set colEndereco = New Collection
        
        lCodigo = lCodigo + 1
        objCliente.lCodigo = lCodigo
        
        objCliente.sRazaoSocial = Trim(Nome)
        
        Select Case TIPO_ASSOC
            
            Case "A"
                objCliente.iTipo = 1
            Case "P"
                objCliente.iTipo = 2
            Case "C"
                objCliente.iTipo = 3
        
        End Select
        
        objCliente.sObservacao = Trim(OBS_1)
        objCliente.dLimiteCredito = 0
        objCliente.dDesconto = 0
        objCliente.iTabelaPreco = 0
        
        'nao incluir cnodicao de pagto
        objCliente.iCondicaoPagto = 0
        objCliente.iAtivo = IIf(STAT_ASSOC = "C", 0, 1)
        objCliente.iAtivoFilial = objCliente.iAtivo
        objCliente.iCodMensagem = 0
        objCliente.lNumPagamentos = 0
        objCliente.iCodPadraoCobranca = 0
        objCliente.iProxCodFilial = 2
        
'???
        sAux = Trim(CPF)
        iPosLixo = InStr(sAux, ".")
        If iPosLixo > 1 Then
            sAux = left(sAux, iPosLixo - 1)
        End If
        objCliente.sCgc = Format(sAux, "00000000000")
            
        objCliente.sInscricaoEstadual = ""
        objCliente.sInscricaoMunicipal = ""
        objCliente.sObservacao = OBS_1
        objCliente.sObservacao2 = OBS_2
        objCliente.iVendedor = 0
        objCliente.iRegiao = 0
        objCliente.sContaContabil = ""
        objCliente.dComissaoVendas = 0
        objCliente.iCodCobrador = 0
        objCliente.iFreqVisitas = 0
        objCliente.dtDataUltVisita = DATA_NULA
        objCliente.iCodTransportadora = 0
        objCliente.iTipoFrete = TIPO_CIF
        
        sAux = Trim(IDENTIDADE)
        If sAux = "0.00" Then sAux = ""
        iPosLixo = InStr(1, sAux, ".00")
        If iPosLixo > 1 Then
            sAux = left(sAux, iPosLixo - 1)
        End If
        objCliente.sRG = sAux
'???
        'Preenche colEndereco
        Set objEndereco = New ClassEndereco
        
        objEndereco.iCodigoPais = 1
        objEndereco.sBairro = Trim(BAIRRO)
        objEndereco.sCEP = Trim(CEP)
        
        'Verifica se o tamanho da string Cidade lida é maior que o permitido
        objEndereco.sCidade = Trim(CIDADE)
        
        objEndereco.sContato = ""
        objEndereco.sEmail = Trim(EMAIL)
        objEndereco.sEndereco = Trim(ENDERECO)
        objEndereco.sFax = ""
        objEndereco.sSiglaEstado = Trim(UF)
        
        'Verifica se o tamanho da string Telefone lida é maior que o permitido
        objEndereco.sTelefone1 = Trim(TELEFONES)
        objEndereco.sTelefone2 = ""
        
        colEndereco.Add objEndereco
        
        Set objEndereco = New ClassEndereco
        
        objEndereco.iCodigoPais = 1
        objEndereco.sBairro = ""
        objEndereco.sCEP = ""
        objEndereco.sCidade = ""
        objEndereco.sContato = ""
        objEndereco.sEmail = ""
        objEndereco.sEndereco = ""
        objEndereco.sFax = ""
        objEndereco.sSiglaEstado = Trim(UF)
        objEndereco.sTelefone1 = ""
        objEndereco.sTelefone2 = ""
        
        colEndereco.Add objEndereco
        
        Set objEndereco = New ClassEndereco
        
        objEndereco.iCodigoPais = 1
        objEndereco.sBairro = ""
        objEndereco.sCEP = ""
        objEndereco.sCidade = ""
        objEndereco.sContato = ""
        objEndereco.sEmail = ""
        objEndereco.sEndereco = ""
        objEndereco.sFax = ""
        objEndereco.sSiglaEstado = Trim(UF)
        objEndereco.sTelefone1 = ""
        objEndereco.sTelefone2 = ""
        
        colEndereco.Add objEndereco
        
        Set objClientesAF = New ClassClientesAF
        Set objCliente.objInfoUsu = objClientesAF
        
        dvmatprinc = MATR_PRINC Mod 10
        lMatriculaPrinc = (MATR_PRINC - dvmatprinc) / 10
        
        With objClientesAF
            .iTipoAssociado = objCliente.iTipo
            .lMatriculaPrinc = lMatriculaPrinc
            .iStatusAssociado = IIf(STAT_ASSOC = "C", 0, 1)
            .lMatriculaSec = CLng(MATR_SEC)
            .lMatriculaAF = CLng(MATR_AF)
            .lMatriculaFRG = CLng(MATR_FRG)
            sAux = Trim(EMPRESA_1)
            .iEmpresa1 = IIf(sAux = "F", 3, IIf(sAux = "R", 4, 2))
            sAux = Trim(EMPRESA_2)
            .iEmpresa2 = IIf(sAux = "F", 3, IIf(sAux = "R", 4, IIf(sAux = "E", 2, 0)))
            .sSexo = Trim(SEXO)
            .sRGOrgaoEmissor = Trim(ORG_EMISS)
            .dtDataExpedicaoRG = StrParaDate(DT_EXPED)
            .dtDataNascimento = DT_NASC
            .dtDataInscricao = DT_INSCR
            .sLocalTrabalho = LOC_TRAB
            .dtDataAdmissaoFurnas = DT_ADMIT_FURNAS
            .sCargo = Trim(CARGO)
            .dtDataAposINSS = StrParaDate(DT_APOS_INSS)
            .dtDataAposFRG = DT_APOSFRG
            
            Select Case Trim(TIPO_APOS)
                Case "E"
                    .iTipoApos = 1
                Case "I"
                    .iTipoApos = 2
                Case "N"
                    .iTipoApos = 3
                Case "T"
                    .iTipoApos = 4
                Case "D"
                    .iTipoApos = 5
            End Select
            
            .dtDataConBenf = StrParaDate(DT_CONBENF)
            .sNoBenef = Trim(NO_BENEF)
            .dtDataFalecimento = DT_FALEC
            .sContrSoc = Trim(CONTR_SOC)
            .sObservacao1 = Trim(OBS_1)
            .sObservacao2 = Trim(OBS_2)
            .iBenemerito = 0
            .iFundador = 0
            .dtDataUltAtualizacao = DT_ULTATU
            .iMatriculaPrincDV = dvmatprinc
        End With
        
        objCliente.sNomeReduzido = Trim(EMPRESA_1) & CStr(lMatriculaPrinc) & "." & Trim(Nome)
        iFimNome = InStr(1, objCliente.sNomeReduzido, " ")
        If iFimNome > 1 Then
            objCliente.sNomeReduzido = left(objCliente.sNomeReduzido, iFimNome - 1)
        End If
        
        'Grava o Cliente
        lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
        If lErro <> SUCESSO Then gError 76334
        
        'Busca o proximo registro de ImportCli
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76341
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76343
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Importa_Clientes = SUCESSO
    
    Exit Function
    
Erro_Importa_Clientes:

    Importa_Clientes = gErr
    
    Select Case gErr
    
        Case 76329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76330, 76331, 76341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCLI", gErr)
        
        Case 76332, 76333, 76334, 76335, 76336
            'Erros tratados nas rotinas chamadas
            
        Case 76342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161823)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function



Private Function FilialCliente_PreencheInfoImportacao(ByVal objFilialCliente As Object, tImportCli As typeImportCli) As Long

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_FilialCliente_PreencheInfoImportacao

    If Len(Trim(tImportCli.sFilialCGC)) > 11 And Len(Trim(tImportCli.sFilialCGC)) <> 14 Then
        objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000000")
    Else
        If Len(Trim(tImportCli.sFilialCGC)) > 8 And Len(Trim(tImportCli.sFilialCGC)) <> 11 Then
            objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000")
        Else
            objFilialCliente.sCgc = tImportCli.sFilialCGC
        End If
    End If
        
    objFilialCliente.sInscricaoEstadual = tImportCli.sFilialInscEstadual
    objFilialCliente.sInscricaoMunicipal = tImportCli.sFilialInscMunicipal
    objFilialCliente.sObservacao = tImportCli.sObservacao
    objFilialCliente.iVendedor = tImportCli.iFilialVendedor
    objFilialCliente.iRegiao = tImportCli.iFilialRegiao
    objFilialCliente.sContaContabil = tImportCli.sFilialContaContabil
    objFilialCliente.dComissaoVendas = tImportCli.dFilialComissaoVendas
    objFilialCliente.iCodCobrador = tImportCli.iFilialCodCobrador
    objFilialCliente.iFreqVisitas = tImportCli.iFilialFreqVisitas
    objFilialCliente.dtDataUltVisita = tImportCli.dtFilialDataUltVisita
    objFilialCliente.iCodTransportadora = tImportCli.iFilialCodTransportadora
    
    If tImportCli.sFilialTipoFrete = "F" Then
        objFilialCliente.iTipoFrete = TIPO_FOB
    ElseIf tImportCli.sFilialTipoFrete = "C" Then
        objFilialCliente.iTipoFrete = TIPO_CIF
    End If
    
    FilialCliente_PreencheInfoImportacao = SUCESSO
     
    Exit Function
    
Erro_FilialCliente_PreencheInfoImportacao:

    FilialCliente_PreencheInfoImportacao = gErr
     
    Select Case Err
          
        Case 76339, 76340, 76385
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161827)
     
    End Select
     
    Exit Function

End Function

Function Importa_Cidades()

Dim lErro As Long, iIndice As Integer, lTransacao As Long
Dim alComando(1 To 3) As Long, sDescricao As String, lQtde As Long

On Error GoTo Erro_Importa_Cidades

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    sDescricao = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Descricao, SUM(Quantidade) FROM ImportCidades GROUP BY Descricao ORDER BY SUM(Quantidade) DESC", sDescricao, lQtde)
    If lErro <> AD_SQL_SUCESSO Then gError 999
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    iIndice = 0
    
    Do While lErro = AD_SQL_SUCESSO
    
        iIndice = iIndice + 1
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO Cidades (Codigo,Descricao) VALUES (?,?)", iIndice, (sDescricao))
        If lErro <> AD_SQL_SUCESSO Then gError 999
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    Loop
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125834

    Importa_Cidades = SUCESSO
     
    Exit Function
    
Erro_Importa_Cidades:

    Importa_Cidades = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161818)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function RelRecibo_Prepara(ByVal lNumIntBaixa As Long) As Long

Dim lErro As Long
Dim objBaixaReceber As New ClassBaixaReceber, lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_RelRecibo_Prepara

    objBaixaReceber.lNumIntBaixa = lNumIntBaixa

    'Lê  a Baixa
    lErro = CF("BaixaRec_Le", objBaixaReceber)
    If lErro <> SUCESSO And lErro <> 46234 Then gError 58967

    If lErro <> SUCESSO Then gError 58968 'Não encontrou

    lErro = CF("RelReciboContabil_Prepara", objBaixaReceber.lNumMovCta, lNumIntRel)
    If lErro <> SUCESSO Then gError 184446
    
    lErro = objRelatorio.ExecutarDireto("Recibos Com Contabilização", "", 1, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 184447
    
    RelRecibo_Prepara = SUCESSO
    
    Exit Function
    
Erro_RelRecibo_Prepara:

    RelRecibo_Prepara = gErr

    Select Case gErr

        Case 58967, 184446, 184447
        
        Case 58968
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAREC_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184455)

    End Select
    
    Exit Function

End Function
