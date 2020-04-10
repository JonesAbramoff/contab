VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABCobrRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'???? acertar a funcao abaixo: tratar erros
Public Function CobrancaEletronica_Abre_TelaGeracaoArq(ByVal sNomeArqParam As String, ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long

Dim lErro As Long

On Error GoTo Erro_CobrancaEletronica_Abre_TelaGeracaoArq

    GeracaoArqRemCobr2.sNomeArqParam = sNomeArqParam
    
    Set GeracaoArqRemCobr2.gobjCobrancaEletronica = objCobrancaEletronica

    GeracaoArqRemCobr2.Show vbModal

    Set GeracaoArqRemCobr2 = Nothing

    CobrancaEletronica_Abre_TelaGeracaoArq = SUCESSO
    
    Exit Function

Erro_CobrancaEletronica_Abre_TelaGeracaoArq:

    CobrancaEletronica_Abre_TelaGeracaoArq = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147237)
    End Select
    
''    Call Rotina_ErrosBatch2("Processamento da Remessa da Cobrança")
    
    Exit Function

End Function

Function CobrancaEletronica_Criar_ArquivoRemessa(ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long

Dim lErro As Long

On Error GoTo Erro_CobrancaEletronica_Criar_ArquivoRemessa

    lErro = CF("CobradorInfo_Le", objCobrancaEletronica.iCobrador, objCobrancaEletronica.objCobradorCNABInfo.colInformacoes)
    If lErro <> SUCESSO Then gError 62008
    
    lErro = CF("CarteiraCobradorInfo_Le_Todas", objCobrancaEletronica.iCobrador, objCobrancaEletronica.colCarteirasCNABInfo)
    If lErro <> SUCESSO Then gError 51681
    
    Select Case objCobrancaEletronica.objCobrador.iCodBanco
        
'        Case 1
'            lErro = CF("Criar_Arquivo_Remessa_BcoBrasil", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62071
    
        Case 230
            lErro = CF("Criar_Arquivo_Remessa_Bandeirantes", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 11111
        
        Case 237
            lErro = CF("Criar_Arquivo_Remessa_Bradesco", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 62070
            
        Case 266
            lErro = CF("Criar_Arquivo_Remessa_Cedula", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 62071
            
'        Case 275
'            lErro = CF("Criar_Arquivo_Remessa_Real", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62071
            
        Case 320
            lErro = CF("Criar_Arquivo_Remessa_BicBanco", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 62071
            
'        Case 341
'            lErro = CF("Criar_Arquivo_Remessa_Itau", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62069
            
        Case 389
            lErro = CF("Criar_Arquivo_Remessa_Mercantil", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 62069
        
'        Case 409
'            lErro = CF("Criar_Arquivo_Remessa_Unibanco", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62068
'
'        Case 422
'            lErro = CF("Criar_Arquivo_Remessa_BancoSafra", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62068
        
'        Case 453
'            lErro = CF("Criar_Arquivo_Remessa_Rural", objCobrancaEletronica)
'            If lErro <> SUCESSO Then gError 62068
        
        Case 604
            lErro = CF("Criar_Arquivo_Remessa_BancoIndustrial", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 86414
        
        Case 719
            lErro = CF("Criar_Arquivo_Remessa_BancoPrimus", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 86414
        
        'cnab 240
        '479: boston, 104:caixa economica federal, 356: real abn amro, 409: unibanco
        '399: HSBC, 341: Itau, 745: Citibank
        '353 / 008 / 033 Santander
        '748: Sicred
        '637: Sofisa
        '85:Cecred
        Case 1, 479, 104, 275, 356, 409, 399, 341, 745, 353, 8, 33, 246, 41, 453, 749, 422, 748, 637, 85
            lErro = CF("Criar_Arquivo_Remessa_CNAB240", objCobrancaEletronica)
            If lErro <> SUCESSO Then gError 62071
    
        Case Else
            Error 59354
            
    End Select
    
    lErro = CF("BancosInfo_Diretorio_Atualiza", objCobrancaEletronica.sDiretorio)
    If lErro <> SUCESSO Then gError 62072
    
    CobrancaEletronica_Criar_ArquivoRemessa = SUCESSO
    
    Exit Function

Erro_CobrancaEletronica_Criar_ArquivoRemessa:

    CobrancaEletronica_Criar_ArquivoRemessa = gErr
    
    Select Case gErr
    
        Case 59354
            '??? falta colocar msg de erro
        
        Case 51681, 62008, 62068, 62069, 62070, 62071, 62072, 11111, 86414
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147236)

    End Select
    
    Exit Function

End Function

Function CobrancaEletronica_Obter_OcorrRemParcRec(ByVal lComando As Long, ByVal objBorderoCobranca As ClassBorderoCobranca, ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long

Dim lErro As Long
Dim tOcorrRemParcRec As typeTipoOcorrRemParcRec
Dim tParcelaRec As typeParcelaReceber
Dim tTitulosRec As typeTituloReceber
Dim tEndereco As typeEndereco
Dim objOcorrRemParcRec As ClassOcorrRemParcRec
Dim colOcorrRemParcRec As New Collection
Dim bIgnoraAltVenc As Boolean
Dim colCobradorInfo As New Collection, objCodigoTexto As AdmCodigoNome

On Error GoTo Erro_CobrancaEletronica_Obter_OcorrRemParcRec
    
    tOcorrRemParcRec.sNumTitCobrador = String(STRING_NUMTITCOBRADOR, 0)
    
    'Lê os dados  de OcorrenRemParcRec para o Bordero passado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, OcorrenciasRemParcRec.Cobrador, CodOcorrencia, DataRegistro, Data, TituloVoltaCarteira, NovaDataVcto, Juros, Instrucao1, DiasDeProtesto1, Instrucao2, DiasDeProtesto2, ValorCobrado, NumBordero, OcorrenciasRemParcRec.NumTitCobrador,NumIntParc,NumSeqOcorr,OcorrenciasRemParcRec.NumIntDoc, ParcelasRecTodas.Status FROM OcorrenciasRemParcRec, ParcelasRecTodas WHERE  NumBordero = ? AND OcorrenciasRemParcRec.NumIntParc = ParcelasRecTodas.NumIntDoc ORDER BY OcorrenciasRemParcRec.NumIntDoc", _
        tOcorrRemParcRec.iFilialEmpresa, tOcorrRemParcRec.iCobrador, tOcorrRemParcRec.iCodOcorrencia, tOcorrRemParcRec.dtDataRegistro, tOcorrRemParcRec.dtData, tOcorrRemParcRec.iTituloVoltaCarteira, tOcorrRemParcRec.dtNovaDataVcto, tOcorrRemParcRec.dJuros, tOcorrRemParcRec.iInstrucao1, tOcorrRemParcRec.iDiasDeProtesto1, tOcorrRemParcRec.iInstrucao2, tOcorrRemParcRec.iDiasDeProtesto2, tOcorrRemParcRec.dValorCobrado, tOcorrRemParcRec.lNumBordero, tOcorrRemParcRec.sNumTitCobrador, _
        tOcorrRemParcRec.lNumIntParc, tOcorrRemParcRec.iNumSeqOcorr, tOcorrRemParcRec.lNumIntDoc, tParcelaRec.iStatus, objBorderoCobranca.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then Error 51647

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51648
    'Se não encontrar nenhuma ocorrencia para o bordero --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 51665
    
    lErro = CF("CobradorInfo_Le", objCobrancaEletronica.objCobrador.iCodigo, colCobradorInfo)
    If lErro <> SUCESSO Then gError 134224

    bIgnoraAltVenc = False
    For Each objCodigoTexto In colCobradorInfo
        If objCodigoTexto.iCodigo = 39 Then 'Código que indica se é para Ignorar as instruções de alteração de vencimento
            If objCodigoTexto.sNome = "1" Then bIgnoraAltVenc = True
            Exit For
        End If
    Next
    
    'Inicializa a coleção de Ocorrências
    Set colOcorrRemParcRec = New Collection
    
    Do While lErro = AD_SQL_SUCESSO

        If Not (tOcorrRemParcRec.iCodOcorrencia = 1 And (tParcelaRec.iStatus = STATUS_BAIXADO Or tParcelaRec.iStatus = STATUS_EXCLUIDO)) Then
        
            If tOcorrRemParcRec.iCodOcorrencia <> 6 Or Not bIgnoraAltVenc Then
        
                Set objOcorrRemParcRec = New ClassOcorrRemParcRec
                'Carrega o objOcorrencia com os dados lidos
                Call Transfere_Dados_OcorrRemParcRec(objOcorrRemParcRec, tOcorrRemParcRec)
                
                'Adiciona a ocorrência na coleção de ocorrências
                colOcorrRemParcRec.Add objOcorrRemParcRec
        
            End If
            
        End If
        
        'Busca a proxima Ocorrencia
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51646
    
    Loop
    
    Set objCobrancaEletronica.colOcorrRemParcRec = colOcorrRemParcRec
    
    CobrancaEletronica_Obter_OcorrRemParcRec = SUCESSO
    
    Exit Function
    
Erro_CobrancaEletronica_Obter_OcorrRemParcRec:

    CobrancaEletronica_Obter_OcorrRemParcRec = Err
    
    Select Case Err
        
        Case 51646, 51647, 51648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OCORRENCIASREMPARCREC", Err)
    
        Case 51665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_SEM_OCORRENCIAS", Err, objBorderoCobranca.lNumBordero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147237)
    
    End Select

    Exit Function

End Function

Private Sub Transfere_Dados_OcorrRemParcRec(objOcorrRemParcRec As ClassOcorrRemParcRec, tOcorrRemParcRec As typeTipoOcorrRemParcRec)
'Trasfere os dados o tOcorrRemParcRec para o objOcorrRemParcRec

    With tOcorrRemParcRec
    
        objOcorrRemParcRec.dJuros = .dJuros
        objOcorrRemParcRec.dtData = .dtData
        objOcorrRemParcRec.dtDataRegistro = .dtDataRegistro
        objOcorrRemParcRec.dtNovaDataVcto = .dtNovaDataVcto
        objOcorrRemParcRec.dValorCobrado = .dValorCobrado
        objOcorrRemParcRec.iCobrador = .iCobrador
        objOcorrRemParcRec.iCodOcorrencia = .iCodOcorrencia
        objOcorrRemParcRec.iDiasDeProtesto1 = .iDiasDeProtesto1
        objOcorrRemParcRec.iDiasDeProtesto2 = .iDiasDeProtesto2
        objOcorrRemParcRec.iFilialEmpresa = .iFilialEmpresa
        objOcorrRemParcRec.iInstrucao1 = .iInstrucao1
        objOcorrRemParcRec.iInstrucao2 = .iInstrucao2
        objOcorrRemParcRec.iNumSeqOcorr = .iNumSeqOcorr
        objOcorrRemParcRec.iTituloVoltaCarteira = .iTituloVoltaCarteira
        objOcorrRemParcRec.lNumBordero = .lNumBordero
        objOcorrRemParcRec.lNumIntParc = .lNumIntParc
        objOcorrRemParcRec.sNumTitCobrador = .sNumTitCobrador
        objOcorrRemParcRec.lNumIntParc = .lNumIntParc
        objOcorrRemParcRec.lNumIntDoc = .lNumIntDoc
    End With
    
    Exit Sub
    
End Sub


Public Function CobrancaEletronica_Obter_DadosParcelaRec(ByVal objParcelaRec As ClassParcelaReceber, ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Obtem os dados para cobrança eletrônica da parcela a receber passada como parâmetro

Dim lErro As Long
Dim lComando As Long, lComando2 As Long
Dim tParcelasRec As typeParcelaReceber
Dim sCGCCedente As String
Dim lNumTitulo As Long
Dim dtDataEmissaoTitulo As Date
Dim lCliente As Long
Dim iFilialCliente As Integer
Dim sCGCSacado As String
Dim sNomeSacado As String
Dim lEndereco As Long
Dim lEnderecoCobranca As Long, sObservacao As String
Dim iEspecie As Integer, sNumTitCobrador As String
Dim dPercMulta As Double, iFilialEmpTitulo As Integer, dPercJurosDiario As Double

On Error GoTo Erro_CobrancaEletronica_Obter_DadosParcelaRec

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51673
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 51673
    
    sCGCCedente = String(STRING_CGC, 0)
    sCGCSacado = String(STRING_CGC, 0)
    sNomeSacado = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    sNumTitCobrador = String(STRING_NUMTITCOBRADOR, 0)
    sObservacao = String(STRING_OBS_PARC_REC, 0)
    
    'Lê os dados nas tabelas FiliaisEmpresa, ParcelasRec, TitulosRec e Filiais Cliente que participarão da cobrança eletrônica
    lErro = Comando_Executar(lComando, "SELECT ParcelasRec.IdImpressaoBoleto, ParcelasRec.NumIntTitulo, ParcelasRec.Observacao, ParcelasRec.NumTitCobrador, FiliaisEmpresa.CGC, ParcelasRec.NumParcela, ParcelasRec.DataVencimento, ParcelasRec.Saldo, ParcelasRec.Aceite, ParcelasRec.Desconto1Ate, ParcelasRec.Desconto1Valor, ParcelasRec.Desconto1Codigo, ParcelasRec.Desconto2Ate, ParcelasRec.Desconto2Valor, ParcelasRec.Desconto2Codigo, ParcelasRec.Desconto3Ate, ParcelasRec.Desconto3Valor, ParcelasRec.Desconto3Codigo, TitulosRec.FilialEmpresa, " & _
        "TitulosRec.NumTitulo, TitulosRec.PercMulta, TitulosRec.PercJurosDiarios, TitulosRec.Especie, TitulosRec.DataEmissao, TitulosRec.Cliente, TitulosRec.Filial, FiliaisClientes.CGC, Clientes.RazaoSocial, FiliaisClientes.EnderecoCobranca,FiliaisClientes.Endereco FROM ParcelasRec, TitulosRec, FiliaisEmpresa, FiliaisClientes, Clientes WHERE ParcelasRec.NumIntTitulo=TitulosRec.NumIntDoc AND FiliaisEmpresa.FilialEmpresa=TitulosRec.FilialEmpresa AND FiliaisClientes.CodCliente=TitulosRec.Cliente AND FiliaisClientes.CodFilial=TitulosRec.Filial AND Clientes.Codigo=TitulosRec.Cliente AND ParcelasRec.NumIntDoc = ?", _
        tParcelasRec.lIdImpressaoBoleto, tParcelasRec.lNumIntTitulo, sObservacao, sNumTitCobrador, sCGCCedente, tParcelasRec.iNumParcela, tParcelasRec.dtDataVencimento, tParcelasRec.dSaldo, tParcelasRec.iAceite, tParcelasRec.dtDesconto1Ate, tParcelasRec.dDesconto1Valor, tParcelasRec.iDesconto1Codigo, tParcelasRec.dtDesconto2Ate, tParcelasRec.dDesconto2Valor, tParcelasRec.iDesconto2Codigo, tParcelasRec.dtDesconto3Ate, tParcelasRec.dDesconto3Valor, tParcelasRec.iDesconto3Codigo, iFilialEmpTitulo, lNumTitulo, dPercMulta, dPercJurosDiario, iEspecie, dtDataEmissaoTitulo, lCliente, iFilialCliente, sCGCSacado, sNomeSacado, lEnderecoCobranca, lEndereco, objParcelaRec.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 51673
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51675
    If lErro <> AD_SQL_SUCESSO Then
    
        'procura nas parcelas de titulos baixados
        lErro = Comando_Executar(lComando2, "SELECT ParcelasRecBaixadas.NumIntTitulo, ParcelasRecBaixadas.Observacao, ParcelasRecBaixadas.NumTitCobrador, FiliaisEmpresa.CGC, ParcelasRecBaixadas.NumParcela, ParcelasRecBaixadas.DataVencimento, 0 as Saldo, ParcelasRecBaixadas.Aceite, ParcelasRecBaixadas.Desconto1Ate, ParcelasRecBaixadas.Desconto1Valor, ParcelasRecBaixadas.Desconto1Codigo, ParcelasRecBaixadas.Desconto2Ate, ParcelasRecBaixadas.Desconto2Valor, ParcelasRecBaixadas.Desconto2Codigo, ParcelasRecBaixadas.Desconto3Ate, ParcelasRecBaixadas.Desconto3Valor, ParcelasRecBaixadas.Desconto3Codigo, TitulosRecBaixados.FilialEmpresa, " & _
            "TitulosRecBaixados.NumTitulo, TitulosRecBaixados.PercMulta, TitulosRecBaixados.PercJurosDiarios, TitulosRecBaixados.Especie, TitulosRecBaixados.DataEmissao, TitulosRecBaixados.Cliente, TitulosRecBaixados.Filial, FiliaisClientes.CGC, Clientes.RazaoSocial, FiliaisClientes.EnderecoCobranca,FiliaisClientes.Endereco FROM ParcelasRecBaixadas, TitulosRecBaixados, FiliaisEmpresa, FiliaisClientes, Clientes WHERE ParcelasRecBaixadas.NumIntTitulo=TitulosRecBaixados.NumIntDoc AND FiliaisEmpresa.FilialEmpresa=TitulosRecBaixados.FilialEmpresa AND FiliaisClientes.CodCliente=TitulosRecBaixados.Cliente AND FiliaisClientes.CodFilial=TitulosRecBaixados.Filial AND Clientes.Codigo=TitulosRecBaixados.Cliente AND ParcelasRecBaixadas.NumIntDoc = ?", _
            tParcelasRec.lNumIntTitulo, sObservacao, sNumTitCobrador, sCGCCedente, tParcelasRec.iNumParcela, tParcelasRec.dtDataVencimento, tParcelasRec.dSaldo, tParcelasRec.iAceite, tParcelasRec.dtDesconto1Ate, tParcelasRec.dDesconto1Valor, tParcelasRec.iDesconto1Codigo, tParcelasRec.dtDesconto2Ate, tParcelasRec.dDesconto2Valor, tParcelasRec.iDesconto2Codigo, tParcelasRec.dtDesconto3Ate, tParcelasRec.dDesconto3Valor, tParcelasRec.iDesconto3Codigo, iFilialEmpTitulo, lNumTitulo, dPercMulta, dPercJurosDiario, iEspecie, dtDataEmissaoTitulo, lCliente, iFilialCliente, sCGCSacado, sNomeSacado, lEnderecoCobranca, lEndereco, objParcelaRec.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 51673
    
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51675
        If lErro <> AD_SQL_SUCESSO Then Error 51674 'Não encontrou
        
    End If
    
    'Preenche o oBJParcela com os dados lidos
    With tParcelasRec
    
        objParcelaRec.iNumParcela = .iNumParcela
        objParcelaRec.dtDataVencimento = .dtDataVencimento
        objParcelaRec.dSaldo = .dSaldo
        objParcelaRec.iAceite = .iAceite
        objParcelaRec.dtDesconto1Ate = .dtDesconto1Ate
        objParcelaRec.dtDesconto2Ate = .dtDesconto2Ate
        objParcelaRec.dtDesconto3Ate = .dtDesconto3Ate
        objParcelaRec.dDesconto1Valor = .dDesconto1Valor
        objParcelaRec.dDesconto2Valor = .dDesconto2Valor
        objParcelaRec.dDesconto3Valor = .dDesconto3Valor
        objParcelaRec.iDesconto1Codigo = .iDesconto1Codigo
        objParcelaRec.iDesconto2Codigo = .iDesconto2Codigo
        objParcelaRec.iDesconto3Codigo = .iDesconto3Codigo
        objParcelaRec.sNumTitCobrador = sNumTitCobrador
        objParcelaRec.sObservacao = sObservacao
        objParcelaRec.lNumIntTitulo = .lNumIntTitulo
        objParcelaRec.lIdImpressaoBoleto = .lIdImpressaoBoleto
        
        objCobrancaEletronica.sCGCCedente = sCGCCedente
        
        objCobrancaEletronica.iFilialEmpTitulo = iFilialEmpTitulo
        objCobrancaEletronica.lNumTitulo = lNumTitulo
        objCobrancaEletronica.iEspecieTitulo = IIf(iEspecie = 0, 1, iEspecie)
        
        objCobrancaEletronica.sCGCSacado = sCGCSacado
        objCobrancaEletronica.sNomeSacado = sNomeSacado
        objCobrancaEletronica.lEnderecoSacado = lEndereco
        objCobrancaEletronica.lEnderecoCobrancaSacado = lEnderecoCobranca
        objCobrancaEletronica.dtDataEmissaoTitulo = dtDataEmissaoTitulo
        objCobrancaEletronica.dPercMulta = dPercMulta
        objCobrancaEletronica.dPercJurosDiario = dPercJurosDiario
        
    End With
    
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec_Cust", objCobrancaEletronica)
    If lErro <> SUCESSO Then gError 51674
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    CobrancaEletronica_Obter_DadosParcelaRec = SUCESSO
    
    Exit Function

Erro_CobrancaEletronica_Obter_DadosParcelaRec:

    CobrancaEletronica_Obter_DadosParcelaRec = Err
    
    Select Case Err
    
        Case 51673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 51673, 51675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", Err)
        
        Case 51674
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, objParcelaRec.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147238)
    
    End Select
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
        
End Function

'########################################################
'Alterado por Wagner
Function CobrancaEletronica_Gera_NomeArquivo(ByVal objCobrador As ClassCobrador, ByVal sDiretorio As String, sNomeArquivo As String, lSequencialArquivo As Long, Optional ByVal iCarteira As Integer = CARTEIRA_SIMPLES) As Long

Dim lErro As Long

On Error GoTo Erro_CobrancaEletronica_Gera_NomeArquivo

    lErro = CF("Geracao_NomeArquivo_CobrancaEletronica", objCobrador, sDiretorio, sNomeArquivo, lSequencialArquivo, iCarteira)
    If lErro <> SUCESSO Then gError 131969
    
    CobrancaEletronica_Gera_NomeArquivo = SUCESSO

    Exit Function

Erro_CobrancaEletronica_Gera_NomeArquivo:

    CobrancaEletronica_Gera_NomeArquivo = gErr
    
    Select Case gErr
    
        Case 131969
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147239)
            
    End Select
    
    Exit Function
            
End Function
'#################################################################

'Maristela(inicio)
'''Public Function CobrancaEletronica_Abre_TelaGeracaoArq(sNomeArqParam As String, objCobrancaEletronica As ClassCobrancaEletronica) As Long
'''
'''Dim lErro As Long
'''
''''On Error GoTo Erro_CobrancaEletronica_Abre_TelaGeracaoArq
'''
'''    lErro = Sistema_Abrir_Batch(sNomeArqParam)
'''    If lErro <> SUCESSO Then Error 11111
'''
'''    Set gcolModulo = New AdmColModulo
'''
'''    lErro = CF("Modulos_Le_Empresa_Filial",glEmpresa, giFilialEmpresa, gcolModulo)
'''    If lErro <> SUCESSO Then Error 22222
'''
'''    Call GeracaoArqRemCobr2.Trata_Parametros(objCobrancaEletronica)
'''
'''    ''Set GeracaoArqRemCobr2.gobjCobrancaEletronica = objCobrancaEletronica
'''
'''    GeracaoArqRemCobr2.Show
'''
'''    Exit Function
'''
'''End Function
'Maristela(inicio)

Public Function CobrancaEletronica_Obter_Endereco(ByVal objCobrancaEletronica As ClassCobrancaEletronica, ByVal objEndereco As ClassEndereco) As Long
'Obtem os dados do endereco para cobrança eletrônica da parcela a receber

Dim lErro As Long

On Error GoTo Erro_CobrancaEletronica_Obter_Endereco

    'Preenche o objEndereço com o endereço de cobrança
    objEndereco.lCodigo = objCobrancaEletronica.lEnderecoCobrancaSacado
    
    'Lê o endereço
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 81708
    If lErro <> SUCESSO Then gError 81709
    
    'se o endereco de cobranca nao estiver preenchido
    If Len(Trim(objEndereco.sEndereco)) = 0 Then
    
        'Preenche o objEndereço com o endereço principal
        objEndereco.lCodigo = objCobrancaEletronica.lEnderecoSacado
    
        'Lê o endereço
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 81710
        If lErro <> SUCESSO Then gError 81711
    
    End If
    
    CobrancaEletronica_Obter_Endereco = SUCESSO
     
    Exit Function
    
Erro_CobrancaEletronica_Obter_Endereco:

    CobrancaEletronica_Obter_Endereco = gErr
     
    Select Case gErr
          
        Case 81708, 81710
        
        Case 81709, 81711
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, objCobrancaEletronica.lCliente, objCobrancaEletronica.iFilialCli)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147240)
     
    End Select
     
    Exit Function

End Function

