VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectBeit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Familias_Le(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tFamilias As typeFamilias
Dim objEnderecoRes As New ClassEndereco
Dim objEnderecoCom As New ClassEndereco
Dim objEnderecoComConj As New ClassEndereco

On Error GoTo Erro_Familias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130410

    'Alocação de espaço no buffer
    tFamilias.sSobrenome = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularNome = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularNomeHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularNomeFirma = String(STRING_RAZAO_SOCIAL, 0)
    tFamilias.sTitularProfissao = String(STRING_PROFISSAO, 0)
    tFamilias.sCohenLeviIsrael = String(1, 0)
    tFamilias.sTitularPai = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularPaiHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularMae = String(STRING_NOME_PESSOA, 0)
    tFamilias.sTitularMaeHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeNome = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeNomeHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeProfissao = String(STRING_PROFISSAO, 0)
    tFamilias.sConjugeNomeFirma = String(STRING_RAZAO_SOCIAL, 0)
    tFamilias.sConjugePai = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugePaiHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeMae = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeMaeHebr = String(STRING_NOME_PESSOA, 0)
    tFamilias.sConjugeSaudacao = String(STRING_FAMILIA_SAUDACAO, 0)
    tFamilias.sTitularSaudacao = String(STRING_FAMILIA_SAUDACAO, 0)

    'Le a tabelaFamilias
    lErro = Comando_Executar(lComando, "SELECT CodFamilia, Sobrenome, TitularNome, TitularNomeHebr, TitularEnderecoRes, TitularNomeFirma, " & _
                "TitularEnderecoCom, LocalCobranca, EstadoCivil, TitularProfissao, TitularDtNasc, TitularDtNascNoite, DataCasamento, " & _
                "DataCasamentoNoite, CohenLeviIsrael, TitularPai, TitularPaiHebr, TitularMae, TitularMaeHebr, TitularDtNascPai, " & _
                "TitularDtNascPaiNoite, TitularDtFalecPai, TitularDtFalecPaiNoite, TitularDtNascMae, TitularDtNascMaeNoite, TitularDtFalecMae, TitularDtFalecMaeNoite, " & _
                "ConjugeNome, ConjugeNomeHebr, ConjugeDtNasc, ConjugeDtNascNoite, ConjugeProfissao, ConjugeNomeFirma, ConjugeEnderecoCom, " & _
                "ConjugePai, ConjugePaiHebr, ConjugeMae, ConjugeMaeHebr, ConjugeDtNascPai, ConjugeDtNascPaiNoite, ConjugeDtFalecPai, " & _
                "ConjugeDtFalecPaiNoite, ConjugeDtNascMae, ConjugeDtNascMaeNoite, ConjugeDtFalecMae, ConjugeDtFalecMaeNoite, ConjugeDtFalec, ConjugeDtFalecNoite, " & _
                "AtualizadoEm, CodCliente, ValorContribuicao, TitularSaudacao, ConjugeSaudacao FROM Familias WHERE CodFamilia= ? ", _
                tFamilias.lCodFamilia, tFamilias.sSobrenome, tFamilias.sTitularNome, tFamilias.sTitularNomeHebr, _
                tFamilias.lTitularEnderecoRes, tFamilias.sTitularNomeFirma, tFamilias.lTitularEnderecoCom, tFamilias.iLocalCobranca, tFamilias.iEstadoCivil, _
                tFamilias.sTitularProfissao, tFamilias.dtTitularDtNasc, tFamilias.iTitularDtNascNoite, tFamilias.dtDataCasamento, tFamilias.iDataCasamentoNoite, _
                tFamilias.sCohenLeviIsrael, tFamilias.sTitularPai, tFamilias.sTitularPaiHebr, tFamilias.sTitularMae, tFamilias.sTitularMaeHebr, _
                tFamilias.dtTitularDtNascPai, tFamilias.iTitularDtNascPaiNoite, tFamilias.dtTitularDtFalecPai, tFamilias.iTitularDtFalecPaiNoite, tFamilias.dtTitularDtNascMae, _
                tFamilias.iTitularDtNascMaeNoite, tFamilias.dtTitularDtFalecMae, tFamilias.iTitularDtFalecMaeNoite, tFamilias.sConjugeNome, tFamilias.sConjugeNomeHebr, _
                tFamilias.dtConjugeDtNasc, tFamilias.iConjugeDtNascNoite, tFamilias.sConjugeProfissao, tFamilias.sConjugeNomeFirma, tFamilias.lConjugeEnderecoCom, _
                tFamilias.sConjugePai, tFamilias.sConjugePaiHebr, tFamilias.sConjugeMae, tFamilias.sConjugeMaeHebr, tFamilias.dtConjugeDtNascPai, _
                tFamilias.iConjugeDtNascPaiNoite, tFamilias.dtConjugeDtFalecPai, tFamilias.iConjugeDtFalecPaiNoite, tFamilias.dtConjugeDtNascMae, tFamilias.iConjugeDtNascMaeNoite, _
                tFamilias.dtConjugeDtFalecMae, tFamilias.iConjugeDtFalecMaeNoite, tFamilias.dtConjugeDtFalec, tFamilias.iConjugeDtFalecNoite, tFamilias.dtAtualizadoEm, _
                tFamilias.lCodCliente, tFamilias.dValorContribuicao, tFamilias.sTitularSaudacao, tFamilias.sConjugeSaudacao, _
                objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 130411

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130412

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 130413

    objFamilias.lCodFamilia = tFamilias.lCodFamilia
    objFamilias.sSobrenome = tFamilias.sSobrenome
    objFamilias.sTitularNome = tFamilias.sTitularNome
    objFamilias.sTitularNomeHebr = tFamilias.sTitularNomeHebr
    objFamilias.lTitularEnderecoRes = tFamilias.lTitularEnderecoRes
    objFamilias.sTitularNomeFirma = tFamilias.sTitularNomeFirma
    objFamilias.lTitularEnderecoCom = tFamilias.lTitularEnderecoCom
    objFamilias.iLocalCobranca = tFamilias.iLocalCobranca
    objFamilias.iEstadoCivil = tFamilias.iEstadoCivil
    objFamilias.sTitularProfissao = tFamilias.sTitularProfissao
    objFamilias.dtTitularDtNasc = tFamilias.dtTitularDtNasc
    objFamilias.iTitularDtNascNoite = tFamilias.iTitularDtNascNoite
    objFamilias.dtDataCasamento = tFamilias.dtDataCasamento
    objFamilias.iDataCasamentoNoite = tFamilias.iDataCasamentoNoite
    objFamilias.sCohenLeviIsrael = tFamilias.sCohenLeviIsrael
    objFamilias.sTitularPai = tFamilias.sTitularPai
    objFamilias.sTitularPaiHebr = tFamilias.sTitularPaiHebr
    objFamilias.sTitularMae = tFamilias.sTitularMae
    objFamilias.sTitularMaeHebr = tFamilias.sTitularMaeHebr
    objFamilias.dtTitularDtNascPai = tFamilias.dtTitularDtNascPai
    objFamilias.iTitularDtNascPaiNoite = tFamilias.iTitularDtNascPaiNoite
    objFamilias.dtTitularDtFalecPai = tFamilias.dtTitularDtFalecPai
    objFamilias.iTitularDtFalecPaiNoite = tFamilias.iTitularDtFalecPaiNoite
    objFamilias.dtTitularDtNascMae = tFamilias.dtTitularDtNascMae
    objFamilias.iTitularDtNascMaeNoite = tFamilias.iTitularDtNascMaeNoite
    objFamilias.dtTitularDtFalecMae = tFamilias.dtTitularDtFalecMae
    objFamilias.iTitularDtFalecMaeNoite = tFamilias.iTitularDtFalecMaeNoite
    objFamilias.sConjugeNome = tFamilias.sConjugeNome
    objFamilias.sConjugeNomeHebr = tFamilias.sConjugeNomeHebr
    objFamilias.dtConjugeDtNasc = tFamilias.dtConjugeDtNasc
    objFamilias.iConjugeDtNascNoite = tFamilias.iConjugeDtNascNoite
    objFamilias.sConjugeProfissao = tFamilias.sConjugeProfissao
    objFamilias.sConjugeNomeFirma = tFamilias.sConjugeNomeFirma
    objFamilias.lConjugeEnderecoCom = tFamilias.lConjugeEnderecoCom
    objFamilias.sConjugePai = tFamilias.sConjugePai
    objFamilias.sConjugePaiHebr = tFamilias.sConjugePaiHebr
    objFamilias.sConjugeMae = tFamilias.sConjugeMae
    objFamilias.sConjugeMaeHebr = tFamilias.sConjugeMaeHebr
    objFamilias.dtConjugeDtNascPai = tFamilias.dtConjugeDtNascPai
    objFamilias.iConjugeDtNascPaiNoite = tFamilias.iConjugeDtNascPaiNoite
    objFamilias.dtConjugeDtFalecPai = tFamilias.dtConjugeDtFalecPai
    objFamilias.iConjugeDtFalecPaiNoite = tFamilias.iConjugeDtFalecPaiNoite
    objFamilias.dtConjugeDtNascMae = tFamilias.dtConjugeDtNascMae
    objFamilias.iConjugeDtNascMaeNoite = tFamilias.iConjugeDtNascMaeNoite
    objFamilias.dtConjugeDtFalecMae = tFamilias.dtConjugeDtFalecMae
    objFamilias.iConjugeDtFalecMaeNoite = tFamilias.iConjugeDtFalecMaeNoite
    objFamilias.dtConjugeDtFalec = tFamilias.dtConjugeDtFalec
    objFamilias.iConjugeDtFalecNoite = tFamilias.iConjugeDtFalecNoite
    objFamilias.dtAtualizadoEm = tFamilias.dtAtualizadoEm
    objFamilias.lCodCliente = tFamilias.lCodCliente
    objFamilias.dValorContribuicao = tFamilias.dValorContribuicao
    objFamilias.sTitularSaudacao = tFamilias.sTitularSaudacao
    objFamilias.sConjugeSaudacao = tFamilias.sConjugeSaudacao

    objEnderecoRes.lCodigo = objFamilias.lTitularEnderecoRes
    
    lErro = CF("Endereco_Le", objEnderecoRes)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 140869
    
    Set objFamilias.objEnderecoRes = objEnderecoRes
    
    objEnderecoCom.lCodigo = objFamilias.lTitularEnderecoCom
    
    lErro = CF("Endereco_Le", objEnderecoCom)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 140870
    
    Set objFamilias.objEnderecoCom = objEnderecoCom
    
    objEnderecoComConj.lCodigo = objFamilias.lConjugeEnderecoCom
    
    lErro = CF("Endereco_Le", objEnderecoComConj)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 140870
    
    Set objFamilias.objEnderecoComConj = objEnderecoComConj
    
    lErro = FilhosFamilias_Le(objFamilias)
    If lErro <> SUCESSO Then gError 140871

    lErro = FamiliasInfo_Le(objFamilias)
    If lErro <> SUCESSO Then gError 140872

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Familias_Le = SUCESSO

    Exit Function

Erro_Familias_Le:

    Familias_Le = gErr

    Select Case gErr

        Case 130410
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130411, 130412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIAS", gErr)

        Case 130413 'Sem dados -> Tratado na rotina chamadora
        
        Case 140869 To 140872

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150284)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function FamiliasInfo_Le(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tFamiliasInfo As typeFamiliasInfo
Dim objFamiliasInfo As New ClassFamiliasInfo

On Error GoTo Erro_FamiliasInfo_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140874

    'Le a tabelaFamiliasInfo
    lErro = Comando_Executar(lComando, "SELECT CodFamilia, Seq, CodInfo, Valor FROM FamiliasInfo WHERE CodFamilia= ? ", _
                tFamiliasInfo.lCodFamilia, tFamiliasInfo.iSeq, tFamiliasInfo.iCodInfo, tFamiliasInfo.iValor, _
                objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 140875

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140876

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFamiliasInfo = New ClassFamiliasInfo

        objFamiliasInfo.lCodFamilia = tFamiliasInfo.lCodFamilia
        objFamiliasInfo.iSeq = tFamiliasInfo.iSeq
        objFamiliasInfo.iCodInfo = tFamiliasInfo.iCodInfo
        objFamiliasInfo.iValor = tFamiliasInfo.iValor
        
        objFamilias.colFamiliaInfo.Add objFamiliasInfo
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140877

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    FamiliasInfo_Le = SUCESSO

    Exit Function

Erro_FamiliasInfo_Le:

    FamiliasInfo_Le = gErr

    Select Case gErr

        Case 140874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140875, 140876, 140877
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIASINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150285)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function FilhosFamilias_Le(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tFilhosFamilias As typeFilhosFamilias
Dim objFilhosFamilias As New ClassFilhosFamilias

On Error GoTo Erro_FilhosFamilias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140878

    'Alocação de espaço no buffer
    tFilhosFamilias.sNome = String(STRING_NOME_PESSOA, 0)
    tFilhosFamilias.sNomeHebr = String(STRING_NOME_PESSOA, 0)
    tFilhosFamilias.sTelefone = String(STRING_TELEFONE, 0)
    tFilhosFamilias.sEmail = String(STRING_EMAIL, 0)

    'Le a tabelaFilhosFamilias
    lErro = Comando_Executar(lComando, "SELECT CodFamilia, SeqFilho, Nome, NomeHebr, DtNasc, DtNascNoite, DtFalecimento, DtFalecimentoNoite, Telefone, Email FROM FilhosFamilias WHERE CodFamilia= ? ", _
                tFilhosFamilias.lCodFamilia, tFilhosFamilias.iSeqFilho, tFilhosFamilias.sNome, tFilhosFamilias.sNomeHebr, _
                tFilhosFamilias.dtDtNasc, tFilhosFamilias.iDtNascNoite, tFilhosFamilias.dtDtFal, tFilhosFamilias.iDtFalNoite, _
                tFilhosFamilias.sTelefone, tFilhosFamilias.sEmail, objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 140879

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140880

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFilhosFamilias = New ClassFilhosFamilias

        objFilhosFamilias.lCodFamilia = tFilhosFamilias.lCodFamilia
        objFilhosFamilias.iSeqFilho = tFilhosFamilias.iSeqFilho
        objFilhosFamilias.sNome = tFilhosFamilias.sNome
        objFilhosFamilias.sNomeHebr = tFilhosFamilias.sNomeHebr
        objFilhosFamilias.dtDtNasc = tFilhosFamilias.dtDtNasc
        objFilhosFamilias.iDtNascNoite = tFilhosFamilias.iDtNascNoite
        objFilhosFamilias.dtDtFal = tFilhosFamilias.dtDtFal
        objFilhosFamilias.iDtFalNoite = tFilhosFamilias.iDtFalNoite
        objFilhosFamilias.sTelefone = tFilhosFamilias.sTelefone
        objFilhosFamilias.sEmail = tFilhosFamilias.sEmail
        
        objFamilias.colFilhos.Add objFilhosFamilias
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140881

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    FilhosFamilias_Le = SUCESSO

    Exit Function

Erro_FilhosFamilias_Le:

    FilhosFamilias_Le = gErr

    Select Case gErr

        Case 140878
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140879, 140880
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILHOSFAMILIAS", gErr)

        Case 140881 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150286)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function FamiliasTipoInfo_Le(ByVal colFamiliasTipoInfo As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tFamiliasTipoInfo As typeFamiliasTipoInfo
Dim objFamiliasTipoInfo As ClassFamiliasTipoInfo

On Error GoTo Erro_FamiliasTipoInfo_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140941

    'Alocação de espaço no buffer
    tFamiliasTipoInfo.sDescricao = String(STRING_FAMILIATIPOINFO_DESCRICAO, 0)
    tFamiliasTipoInfo.sSigla = String(STRING_FAMILIATIPOINFO_SIGLA, 0)

    'Le a tabelaFamiliasTipoInfo
    lErro = Comando_Executar(lComando, "SELECT CodInfo, Descricao, Sigla, ValidoPara, Posicao FROM FamiliasTipoInfo ", _
                tFamiliasTipoInfo.iCodInfo, tFamiliasTipoInfo.sDescricao, tFamiliasTipoInfo.sSigla, tFamiliasTipoInfo.iValidoPara, _
                tFamiliasTipoInfo.iPosicao)
    If lErro <> AD_SQL_SUCESSO Then gError 140942

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140943

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFamiliasTipoInfo = New ClassFamiliasTipoInfo

        objFamiliasTipoInfo.iCodInfo = tFamiliasTipoInfo.iCodInfo
        objFamiliasTipoInfo.sDescricao = tFamiliasTipoInfo.sDescricao
        objFamiliasTipoInfo.sSigla = tFamiliasTipoInfo.sSigla
        objFamiliasTipoInfo.iValidoPara = tFamiliasTipoInfo.iValidoPara
        objFamiliasTipoInfo.iPosicao = tFamiliasTipoInfo.iPosicao
    
        colFamiliasTipoInfo.Add objFamiliasTipoInfo
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140944

    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    FamiliasTipoInfo_Le = SUCESSO

    Exit Function

Erro_FamiliasTipoInfo_Le:

    FamiliasTipoInfo_Le = gErr

    Select Case gErr

        Case 140941
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140942, 140943, 140944
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIASTIPOINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150287)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

