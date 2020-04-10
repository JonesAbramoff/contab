VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Cliente_Le_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tClientesAF As typeClientesAF
Dim objClientesAF As New ClassClientesAF

On Error GoTo Erro_Cliente_Le_Customizado

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192936

    'Alocação de espaço no buffer
    tClientesAF.sSexo = String(STRING_CLIENTESAF_SEXO, 0)
    tClientesAF.sRGOrgaoEmissor = String(STRING_CLIENTESAF_RGORGAOEMISSOR, 0)
    tClientesAF.sLocalTrabalho = String(STRING_CLIENTESAF_LOCALTRABALHO, 0)
    tClientesAF.sCargo = String(STRING_CLIENTESAF_CARGO, 0)
    tClientesAF.sNoBenef = String(STRING_CLIENTESAF_NOBENEF, 0)
    tClientesAF.sContrSoc = String(STRING_CLIENTESAF_CONTRSOC, 0)
    tClientesAF.sObservacao1 = String(STRING_CLIENTESAF_OBS1, 0)
    tClientesAF.sObservacao2 = String(STRING_CLIENTESAF_OBS2, 0)

    'Le a tabelaClientesAF
    lErro = Comando_Executar(lComando, "SELECT Cliente, TipoAssociado, MatriculaPrinc, StatusAssociado, MatriculaSec, MatriculaAF, " & _
                "MatriculaFRG, Empresa1, Empresa2, Sexo, RGOrgaoEmissor, DataExpedicaoRG, DataNascimento, " & _
                "DataInscricao, LocalTrabalho, DataAdmissaoFurnas, Cargo, DataAposINSS, DataAposFRG, TipoApos, " & _
                "DataConBenf, NoBenef, DataFalecimento, ContrSoc, Observacao1, Observacao2, Benemerito, " & _
                "Fundador, DataUltAtualizacao, MatriculaPrincDV FROM ClientesAF WHERE Cliente= ? ", _
                tClientesAF.lCliente, tClientesAF.iTipoAssociado, tClientesAF.lMatriculaPrinc, tClientesAF.iStatusAssociado, _
                tClientesAF.lMatriculaSec, tClientesAF.lMatriculaAF, tClientesAF.lMatriculaFRG, tClientesAF.iEmpresa1, tClientesAF.iEmpresa2, _
                tClientesAF.sSexo, tClientesAF.sRGOrgaoEmissor, tClientesAF.dtDataExpedicaoRG, tClientesAF.dtDataNascimento, tClientesAF.dtDataInscricao, _
                tClientesAF.sLocalTrabalho, tClientesAF.dtDataAdmissaoFurnas, tClientesAF.sCargo, tClientesAF.dtDataAposINSS, tClientesAF.dtDataAposFRG, _
                tClientesAF.iTipoApos, tClientesAF.dtDataConBenf, tClientesAF.sNoBenef, tClientesAF.dtDataFalecimento, tClientesAF.sContrSoc, _
                tClientesAF.sObservacao1, tClientesAF.sObservacao2, tClientesAF.iBenemerito, tClientesAF.iFundador, tClientesAF.dtDataUltAtualizacao, tClientesAF.iMatriculaPrincDV, _
                objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192937

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192938

    'Sem Dados
    'If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    If lErro <> AD_SQL_SEM_DADOS Then

        objClientesAF.lCliente = tClientesAF.lCliente
        objClientesAF.iTipoAssociado = tClientesAF.iTipoAssociado
        objClientesAF.lMatriculaPrinc = tClientesAF.lMatriculaPrinc
        objClientesAF.iStatusAssociado = tClientesAF.iStatusAssociado
        objClientesAF.lMatriculaSec = tClientesAF.lMatriculaSec
        objClientesAF.lMatriculaAF = tClientesAF.lMatriculaAF
        objClientesAF.lMatriculaFRG = tClientesAF.lMatriculaFRG
        objClientesAF.iEmpresa1 = tClientesAF.iEmpresa1
        objClientesAF.iEmpresa2 = tClientesAF.iEmpresa2
        objClientesAF.sSexo = tClientesAF.sSexo
        objClientesAF.sRGOrgaoEmissor = tClientesAF.sRGOrgaoEmissor
        objClientesAF.dtDataExpedicaoRG = tClientesAF.dtDataExpedicaoRG
        objClientesAF.dtDataNascimento = tClientesAF.dtDataNascimento
        objClientesAF.dtDataInscricao = tClientesAF.dtDataInscricao
        objClientesAF.sLocalTrabalho = tClientesAF.sLocalTrabalho
        objClientesAF.dtDataAdmissaoFurnas = tClientesAF.dtDataAdmissaoFurnas
        objClientesAF.sCargo = tClientesAF.sCargo
        objClientesAF.dtDataAposINSS = tClientesAF.dtDataAposINSS
        objClientesAF.dtDataAposFRG = tClientesAF.dtDataAposFRG
        objClientesAF.iTipoApos = tClientesAF.iTipoApos
        objClientesAF.dtDataConBenf = tClientesAF.dtDataConBenf
        objClientesAF.sNoBenef = tClientesAF.sNoBenef
        objClientesAF.dtDataFalecimento = tClientesAF.dtDataFalecimento
        objClientesAF.sContrSoc = tClientesAF.sContrSoc
        objClientesAF.sObservacao1 = tClientesAF.sObservacao1
        objClientesAF.sObservacao2 = tClientesAF.sObservacao2
        objClientesAF.iBenemerito = tClientesAF.iBenemerito
        objClientesAF.iFundador = tClientesAF.iFundador
        objClientesAF.dtDataUltAtualizacao = tClientesAF.dtDataUltAtualizacao
        objClientesAF.iMatriculaPrincDV = tClientesAF.iMatriculaPrincDV
        
    End If
    
    Set objCliente.objInfoUsu = objClientesAF

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Cliente_Le_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Le_Customizado:

    Cliente_Le_Customizado = gErr

    Select Case gErr

        Case 192936
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192937, 192938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192939)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cliente_Obtem_NomeBrowse(sNomeBrowse As String) As Long
    sNomeBrowse = "ClientesFiliais_AFLista"
    Cliente_Obtem_NomeBrowse = SUCESSO
End Function
