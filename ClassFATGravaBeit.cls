VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaBeit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Familias_Grava(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Familias_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130423

    lErro = CF("Familias_Grava_EmTrans", objFamilias)
    If lErro <> SUCESSO Then gError 130424

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130425

    Familias_Grava = SUCESSO

    Exit Function

Erro_Familias_Grava:

    Familias_Grava = gErr

    Select Case gErr

        Case 130423
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 130424

        Case 130425
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150044)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Familias_Grava_EmTrans(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lCodEndereco As Long
Dim lEnderecoRes As Long
Dim lEnderecoCom As Long
Dim lEnderecoComConj As Long
Dim sSQL As String

On Error GoTo Erro_Familias_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130426
    Next
    
    If gsUsuario <> "Nadia" And InStr(1, "Leblon/Lar/Cheder/Templo/Six/Guemach", gsNomeFilialEmpresa) <> 0 And gsNomeEmpresa = "Beit Lubavitch" Then gError 130430
    
    'Le a tabelaFamilias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ConjugeEnderecoCom, TitularEnderecoCom, TitularEnderecoRes FROM Familias WHERE CodFamilia= ? ", _
                0, lEnderecoComConj, lEnderecoCom, lEnderecoRes, objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 130427

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130428

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 130429
      
        lErro = FilhosFamilias_Exclui(objFamilias)
        If lErro <> SUCESSO Then gError 140890

        lErro = FamiliasInfo_Exclui(objFamilias)
        If lErro <> SUCESSO Then gError 140891

        objFamilias.objEnderecoCom.lCodigo = lEnderecoCom

        lErro = CF("Endereco_Grava", objFamilias.objEnderecoCom)
        If lErro <> SUCESSO Then gError 140908

        objFamilias.objEnderecoRes.lCodigo = lEnderecoRes

        lErro = CF("Endereco_Grava", objFamilias.objEnderecoRes)
        If lErro <> SUCESSO Then gError 140909
        
        objFamilias.objEnderecoComConj.lCodigo = lEnderecoComConj
        
        lErro = CF("Endereco_Grava", objFamilias.objEnderecoComConj)
        If lErro <> SUCESSO Then gError 140909
       
        objFamilias.lTitularEnderecoRes = objFamilias.objEnderecoRes.lCodigo
        objFamilias.lTitularEnderecoCom = objFamilias.objEnderecoCom.lCodigo
        objFamilias.lConjugeEnderecoCom = objFamilias.objEnderecoComConj.lCodigo
       
        sSQL = "UPDATE Familias SET Sobrenome= ? , TitularNome= ? , TitularNomeHebr= ? , TitularEnderecoRes= ? , " & _
                "TitularNomeFirma= ? , TitularEnderecoCom= ? , LocalCobranca= ? , EstadoCivil= ? , TitularProfissao= ? , " & _
                "TitularDtNasc= ? , TitularDtNascNoite= ? , TitularDtNascHebr= ? ,DataCasamento= ? , DataCasamentoNoite= ? , DataCasamentoHebr= ? , CohenLeviIsrael= ? , " & _
                "TitularPai= ? , TitularPaiHebr= ? , TitularMae= ? , TitularMaeHebr= ? , TitularDtNascPai= ? , " & _
                "TitularDtNascPaiNoite= ? , TitularDtNascPaiHebr = ?, TitularDtFalecPai= ? , TitularDtFalecPaiNoite= ? , TitularDtFalecPaiHebr= ? , TitularDtNascMae= ? , TitularDtNascMaeNoite= ? , TitularDtNascMaeHebr= ? , " & _
                "TitularDtFalecMae= ? , TitularDtFalecMaeNoite= ? , TitularDtFalecMaeHebr= ? , ConjugeNome= ? , ConjugeNomeHebr= ? , ConjugeDtNasc= ? , " & _
                "ConjugeDtNascNoite= ? , ConjugeDtNascHebr= ? , ConjugeProfissao= ? , ConjugeNomeFirma= ? , ConjugeEnderecoCom= ? , ConjugePai= ? , " & _
                "ConjugePaiHebr= ? , ConjugeMae= ? , ConjugeMaeHebr= ? , ConjugeDtNascPai= ? , ConjugeDtNascPaiNoite= ? , ConjugeDtNascPaiHebr= ? , " & _
                "ConjugeDtFalecPai= ? , ConjugeDtFalecPaiNoite= ? , ConjugeDtFalecPaiHebr= ? , ConjugeDtNascMae= ? , ConjugeDtNascMaeNoite= ? , ConjugeDtNascMaeHebr= ? , ConjugeDtFalecMae= ? , " & _
                "ConjugeDtFalecMaeNoite= ? , ConjugeDtFalecMaeHebr= ? , ConjugeDtFalec= ? , ConjugeDtFalecNoite= ? , ConjugeDtFalecHebr= ? , AtualizadoEm= ? , CodCliente= ? , " & _
                "ValorContribuicao= ?, TitularSaudacao = ?, ConjugeSaudacao = ?, Usuario = ?, " & _
                "AN=?,CD=?,CH=?,CJ=?,H=?,H1=?,H2=?,LE=?,LR=?,PA=?,RE=?,SH=?,SI=?,TH=?,VF=?,ANConj=?,CDConj=?,CHConj=?,CJConj=?,HConj=?," & _
                "H1Conj=?, H2Conj=?, LEConj=?, LRConj=?, PAConj=?, REConj=?, SHConj=?, SIConj=?, THConj=?, VFConj=?"
        
        With objFamilias
            lErro = Comando_ExecutarPos(alComando(1), sSQL, alComando(0), _
                .sSobrenome, .sTitularNome, .sTitularNomeHebr, .lTitularEnderecoRes, _
                .sTitularNomeFirma, .lTitularEnderecoCom, .iLocalCobranca, .iEstadoCivil, .sTitularProfissao, _
                .dtTitularDtNasc, .iTitularDtNascNoite, .sTitularDtNascHebr, .dtDataCasamento, .iDataCasamentoNoite, .sDataCasamentoHebr, .sCohenLeviIsrael, _
                .sTitularPai, .sTitularPaiHebr, .sTitularMae, .sTitularMaeHebr, .dtTitularDtNascPai, _
                .iTitularDtNascPaiNoite, .sTitularDtNascPaiHebr, .dtTitularDtFalecPai, .iTitularDtFalecPaiNoite, .sTitularDtFalecPaiHebr, .dtTitularDtNascMae, .iTitularDtNascMaeNoite, .sTitularDtNascMaeHebr, _
                .dtTitularDtFalecMae, .iTitularDtFalecMaeNoite, .sTitularDtFalecMaeHebr, .sConjugeNome, .sConjugeNomeHebr, .dtConjugeDtNasc, _
                .iConjugeDtNascNoite, .sConjugeDtNascHebr, .sConjugeProfissao, .sConjugeNomeFirma, .lConjugeEnderecoCom, .sConjugePai, _
                .sConjugePaiHebr, .sConjugeMae, .sConjugeMaeHebr, .dtConjugeDtNascPai, .iConjugeDtNascPaiNoite, .sConjugeDtNascPaiHebr, _
                .dtConjugeDtFalecPai, .iConjugeDtFalecPaiNoite, .sConjugeDtFalecPaiHebr, .dtConjugeDtNascMae, .iConjugeDtNascMaeNoite, .sConjugeDtNascMaeHebr, .dtConjugeDtFalecMae, _
                .iConjugeDtFalecMaeNoite, .sConjugeDtFalecMaeHebr, .dtConjugeDtFalec, .iConjugeDtFalecNoite, .sConjugeDtFalecHebr, .dtAtualizadoEm, .lCodCliente, _
                .dValorContribuicao, .sTitularSaudacao, .sConjugeSaudacao, gsUsuario, .iAN, .iCD, .iCH, .iCJ, .iH, .iH1, .iH2, .iLE, .iLR, .iPA, .iRE, .iSH, .iSI, .iTH, .iVF, .iANConj, .iCDConj, .iCHConj, .iCJConj, .iHConj, _
                .iH1Conj, .iH2Conj, .iLEConj, .iLRConj, .iPAConj, .iREConj, .iSHConj, .iSIConj, .iTHConj, .iVFConj)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130430
         
    Else

        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 140904

        objFamilias.objEnderecoCom.lCodigo = lCodEndereco

        lErro = CF("Endereco_Grava", objFamilias.objEnderecoCom)
        If lErro <> SUCESSO Then gError 140905

        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 140906

        objFamilias.objEnderecoRes.lCodigo = lCodEndereco

        lErro = CF("Endereco_Grava", objFamilias.objEnderecoRes)
        If lErro <> SUCESSO Then gError 140907

        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 140950

        objFamilias.objEnderecoComConj.lCodigo = lCodEndereco

        lErro = CF("Endereco_Grava", objFamilias.objEnderecoComConj)
        If lErro <> SUCESSO Then gError 140951
        
        objFamilias.lTitularEnderecoRes = objFamilias.objEnderecoRes.lCodigo
        objFamilias.lTitularEnderecoCom = objFamilias.objEnderecoCom.lCodigo
        objFamilias.lConjugeEnderecoCom = objFamilias.objEnderecoComConj.lCodigo

        sSQL = "INSERT INTO Familias( CodFamilia, Sobrenome, TitularNome, TitularNomeHebr, " & _
                "TitularEnderecoRes, TitularNomeFirma, TitularEnderecoCom, LocalCobranca, EstadoCivil, " & _
                "TitularProfissao, TitularDtNasc, TitularDtNascNoite, TitularDtNascHebr, DataCasamento, DataCasamentoNoite, DataCasamentoHebr, " & _
                "CohenLeviIsrael, TitularPai, TitularPaiHebr, TitularMae, TitularMaeHebr, " & _
                "TitularDtNascPai, TitularDtNascPaiNoite, TitularDtNascPaiHebr, TitularDtFalecPai, TitularDtFalecPaiNoite, TitularDtFalecPaiHebr, TitularDtNascMae, " & _
                "TitularDtNascMaeNoite, TitularDtNascMaeHebr, TitularDtFalecMae, TitularDtFalecMaeNoite, TitularDtFalecMaeHebr, ConjugeNome, ConjugeNomeHebr, " & _
                "ConjugeDtNasc, ConjugeDtNascNoite, ConjugeDtNascHebr, ConjugeProfissao, ConjugeNomeFirma, ConjugeEnderecoCom, " & _
                "ConjugePai, ConjugePaiHebr, ConjugeMae, ConjugeMaeHebr, ConjugeDtNascPai, " & _
                "ConjugeDtNascPaiNoite, ConjugeDtNascPaiHebr, ConjugeDtFalecPai, ConjugeDtFalecPaiNoite, ConjugeDtFalecPaiHebr, ConjugeDtNascMae, ConjugeDtNascMaeNoite, ConjugeDtNascMaeHebr, " & _
                "ConjugeDtFalecMae, ConjugeDtFalecMaeNoite, ConjugeDtFalecMaeHebr, ConjugeDtFalec, ConjugeDtFalecNoite, ConjugeDtFalecHebr, AtualizadoEm, " & _
                "CodCliente, ValorContribuicao, TitularSaudacao, ConjugeSaudacao, Usuario, " & _
                "AN,CD,CH,CJ,H,H1,H2,LE,LR,PA,RE,SH,SI,TH,VF,ANConj,CDConj,CHConj,CJConj,HConj," & _
                "H1Conj , H2Conj, LEConj, LRConj, PAConj, REConj, SHConj, SIConj, THConj, VFConj)" & _
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        With objFamilias
            lErro = Comando_Executar(alComando(1), sSQL, .lCodFamilia, .sSobrenome, .sTitularNome, .sTitularNomeHebr, _
                .lTitularEnderecoRes, .sTitularNomeFirma, .lTitularEnderecoCom, .iLocalCobranca, .iEstadoCivil, _
                .sTitularProfissao, .dtTitularDtNasc, .iTitularDtNascNoite, .sTitularDtNascHebr, .dtDataCasamento, .iDataCasamentoNoite, .sDataCasamentoHebr, _
                .sCohenLeviIsrael, .sTitularPai, .sTitularPaiHebr, .sTitularMae, .sTitularMaeHebr, _
                .dtTitularDtNascPai, .iTitularDtNascPaiNoite, .sTitularDtNascPaiHebr, .dtTitularDtFalecPai, .iTitularDtFalecPaiNoite, .sTitularDtFalecPaiHebr, .dtTitularDtNascMae, _
                .iTitularDtNascMaeNoite, .sTitularDtNascMaeHebr, .dtTitularDtFalecMae, .iTitularDtFalecMaeNoite, .sTitularDtFalecMaeHebr, .sConjugeNome, .sConjugeNomeHebr, _
                .dtConjugeDtNasc, .iConjugeDtNascNoite, .sConjugeDtNascHebr, .sConjugeProfissao, .sConjugeNomeFirma, .lConjugeEnderecoCom, _
                .sConjugePai, .sConjugePaiHebr, .sConjugeMae, .sConjugeMaeHebr, .dtConjugeDtNascPai, _
                .iConjugeDtNascPaiNoite, .sConjugeDtNascPaiHebr, .dtConjugeDtFalecPai, .iConjugeDtFalecPaiNoite, .sConjugeDtFalecPaiHebr, .dtConjugeDtNascMae, .iConjugeDtNascMaeNoite, .sConjugeDtNascMaeHebr, _
                .dtConjugeDtFalecMae, .iConjugeDtFalecMaeNoite, .sConjugeDtFalecMaeHebr, .dtConjugeDtFalec, .iConjugeDtFalecNoite, .sConjugeDtFalecHebr, .dtAtualizadoEm, _
                .lCodCliente, .dValorContribuicao, .sTitularSaudacao, .sConjugeSaudacao, gsUsuario, .iAN, .iCD, .iCH, .iCJ, .iH, .iH1, .iH2, .iLE, .iLR, .iPA, .iRE, .iSH, .iSI, .iTH, .iVF, .iANConj, .iCDConj, .iCHConj, .iCJConj, .iHConj, _
                .iH1Conj, .iH2Conj, .iLEConj, .iLRConj, .iPAConj, .iREConj, .iSHConj, .iSIConj, .iTHConj, .iVFConj)
        End With
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130431
        
    End If

    lErro = FilhosFamilias_Insere(objFamilias)
    If lErro <> SUCESSO Then gError 140892

    lErro = FamiliasInfo_Insere(objFamilias)
    If lErro <> SUCESSO Then gError 140893

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Familias_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Familias_Grava_EmTrans:

    Familias_Grava_EmTrans = gErr

    Select Case gErr

        Case 130426
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130427, 130428
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIAS", gErr)

        Case 130429
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_FAMILIAS", gErr)

        Case 130430
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_FAMILIAS", gErr)

        Case 130431
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FAMILIAS", gErr)
            
        Case 140890 To 140893, 140904 To 140909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150045)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function Familias_Exclui(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Familias_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130414

    lErro = CF("Familias_Exclui_EmTrans", objFamilias)
    If lErro <> SUCESSO Then gError 130415

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130416

    Familias_Exclui = SUCESSO

    Exit Function

Erro_Familias_Exclui:

    Familias_Exclui = gErr

    Select Case gErr

        Case 130414
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 130415

        Case 130416
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150046)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Familias_Exclui_EmTrans(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lEnderecoRes As Long
Dim lEnderecoCom As Long
Dim lEnderecoComConj As Long

On Error GoTo Erro_Familias_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130417
    Next

    If gsUsuario <> "Nadia" And InStr(1, "Leblon/Lar/Cheder/Templo/Six/Guemach", gsNomeFilialEmpresa) <> 0 And gsNomeEmpresa = "Beit Lubavitch" Then gError 130422
    
    'Le a tabelaFamilias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ConjugeEnderecoCom, TitularEnderecoCom, TitularEnderecoRes FROM Familias WHERE CodFamilia= ? ", _
                0, lEnderecoComConj, lEnderecoCom, lEnderecoRes, objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 130418

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130419

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 130420

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 130421

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Familias", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 130422
    
    lErro = FilhosFamilias_Exclui(objFamilias)
    If lErro <> SUCESSO Then gError 140902

    lErro = FamiliasInfo_Exclui(objFamilias)
    If lErro <> SUCESSO Then gError 140903
    
    lErro = CF("Endereco_Exclui", lEnderecoRes)
    If lErro <> SUCESSO Then gError 140910

    lErro = CF("Endereco_Exclui", lEnderecoCom)
    If lErro <> SUCESSO Then gError 140911

    lErro = CF("Endereco_Exclui", lEnderecoComConj)
    If lErro <> SUCESSO Then gError 140911

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Familias_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Familias_Exclui_EmTrans:

    Familias_Exclui_EmTrans = gErr

    Select Case gErr

        Case 130417
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130418, 130419
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIAS", gErr)

        Case 130420
            Call Rotina_Erro(vbOKOnly, "ERRO_FAMILIAS_NAO_CADASTRADO", gErr, objFamilias.lCodFamilia)

        Case 130421
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_FAMILIAS", gErr)

        Case 130422
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FAMILIAS", gErr)
            
        Case 140902, 140903, 140910, 140911

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150047)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function FilhosFamilias_Insere(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim objFilhosFamilias As ClassFilhosFamilias

On Error GoTo Erro_FilhosFamilias_Insere

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140884
    
    For Each objFilhosFamilias In objFamilias.colFilhos
        
        objFilhosFamilias.lCodFamilia = objFamilias.lCodFamilia

        With objFilhosFamilias
            lErro = Comando_Executar(lComando, "INSERT INTO FilhosFamilias( CodFamilia, SeqFilho, Nome, NomeHebr, " & _
                "DtNasc, DtNascNoite, DtNascHebr, dtFalecimento, dtFalecimentoNoite, dtFalecimentoHebr, Telefone, Email, AN,CD,CH,CJ,H,H1,H2,LE,LR,PA,RE,SH,SI,TH,VF) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lCodFamilia, .iSeqFilho, .sNome, .sNomeHebr, _
                .dtDtNasc, .iDtNascNoite, .sDtNascHebr, .dtDtFal, .iDtFalNoite, .sDtFalHebr, .sTelefone, .sEmail, .iAN, .iCD, .iCH, .iCJ, .iH, .iH1, .iH2, .iLE, .iLR, .iPA, .iRE, .iSH, .iSI, .iTH, .iVF)
        End With
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140885

    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    FilhosFamilias_Insere = SUCESSO

    Exit Function

Erro_FilhosFamilias_Insere:

    FilhosFamilias_Insere = gErr

    Select Case gErr

        Case 140884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140885
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILHOSFAMILIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150048)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function FilhosFamilias_Exclui(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_FilhosFamilias_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140886
    Next

    'Le a tabelaFilhosFamilias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM FilhosFamilias WHERE CodFamilia= ? ", _
                0, iAux, objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 140887

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140888

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 140890
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM FilhosFamilias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 140891

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140889
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FilhosFamilias_Exclui = SUCESSO

    Exit Function

Erro_FilhosFamilias_Exclui:

    FilhosFamilias_Exclui = gErr

    Select Case gErr

        Case 140886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140887, 140888, 140889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILHOSFAMILIAS", gErr)
        
        Case 140890
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_FILHOSFAMILIAS", gErr)

        Case 140891
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILHOSFAMILIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150049)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function FamiliasInfo_Insere(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim objFamiliasInfo As ClassFamiliasInfo

On Error GoTo Erro_FamiliasInfo_Insere

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140894

    For Each objFamiliasInfo In objFamilias.colFamiliaInfo
    
        objFamiliasInfo.lCodFamilia = objFamilias.lCodFamilia
    
        lErro = Comando_Executar(lComando, "INSERT INTO FamiliasInfo( CodFamilia, Seq, CodInfo, Valor) VALUES (?,?,?,?)", _
                objFamiliasInfo.lCodFamilia, objFamiliasInfo.iSeq, objFamiliasInfo.iCodInfo, objFamiliasInfo.iValor)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140895

    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    FamiliasInfo_Insere = SUCESSO

    Exit Function

Erro_FamiliasInfo_Insere:

    FamiliasInfo_Insere = gErr

    Select Case gErr

        Case 140894
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140895
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FAMILIASINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150050)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function FamiliasInfo_Exclui(ByVal objFamilias As ClassFamilias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_FamiliasInfo_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140896
    Next

    'Le a tabelaFamiliasInfo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM FamiliasInfo WHERE CodFamilia= ? ", _
                0, iAux, objFamilias.lCodFamilia)
    If lErro <> AD_SQL_SUCESSO Then gError 140897

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140898

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 140899
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM FamiliasInfo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 140900
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140901
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FamiliasInfo_Exclui = SUCESSO

    Exit Function

Erro_FamiliasInfo_Exclui:

    FamiliasInfo_Exclui = gErr

    Select Case gErr

        Case 140896
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140897, 140898, 140901
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FAMILIASINFO", gErr)

        Case 140899
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_FAMILIASINFO", gErr)

        Case 140900
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FAMILIASINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150051)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Familias_Automatico(lCodigo As Long) As Long
'Obtem o próximo número de Familias

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_Familias_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_FAMILIAS", "Familias", "CodFamilia", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 140952

    lCodigo = lNumAuto

    Familias_Automatico = SUCESSO

    Exit Function

Erro_Familias_Automatico:

    Familias_Automatico = gErr

   Select Case gErr

        'Erro já tratado
        Case 140953

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150052)

    End Select

    Exit Function

End Function

Public Function NFiscalContrato_Gera_InfoBoletos(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, sObservacao As String, objItemNF As ClassItemNF, objParcRec As ClassParcelaReceber

On Error GoTo Erro_NFiscalContrato_Gera_InfoBoletos


    sObservacao = ""
    'sObservacao = FormataCpoTexto("PAGÁVEL EM QUALQUER BANCO MESMO APÓS O VENCIMENTO", TAMANHO_CADA_INSTRUCAO_BOLETO)
    
    For Each objItemNF In objNFiscal.ColItensNF
    
        sObservacao = sObservacao & FormataCpoTexto(objItemNF.sDescricaoItem & " - " & UCase(Format(objItemNF.objCobrItensContrato.dtDataRefIni, "MMM/YY")), TAMANHO_CADA_INSTRUCAO_BOLETO)
        
    Next

    For Each objParcRec In objNFiscal.ColParcelaReceber
    
        objParcRec.sObservacao = sObservacao
        
    Next
    
    NFiscalContrato_Gera_InfoBoletos = SUCESSO
    
    Exit Function
    
Erro_NFiscalContrato_Gera_InfoBoletos:

    NFiscalContrato_Gera_InfoBoletos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184418)

    End Select
    
    Exit Function

End Function

Public Function Preenche_Instrucoes_Boleto(ByVal objTituloReceber As ClassTituloReceber, ByVal objParcelaReceber As ClassParcelaReceber, ByVal sSimbolo As String, sInstrucoes As String) As Long

Dim lErro As Long, lComando As Long, sDescricao As String
Dim sInstrucao_MultaJuros As String

On Error GoTo Erro_Preenche_Instrucoes_Boleto

    sInstrucoes = objParcelaReceber.sObservacao
    
    If sInstrucoes = "" Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 140884
    
        sDescricao = String(STRING_ITEMNF_DESCRICAO, 0)
        lErro = Comando_Executar(lComando, "SELECT DescricaoItem FROM ItensNFiscal, NFiscal WHERE ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND NFiscal.ClasseDocCPR = ? AND NFiscal.NumIntDocCPR = ? AND Item = 1", _
            sDescricao, CPR_TITULO_RECEBER, objTituloReceber.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 140885
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140885
        
        If lErro = AD_SQL_SUCESSO Then sInstrucoes = sDescricao
        
        'Fecha Comando
        Call Comando_Fechar(lComando)
    
    End If
    
    sInstrucao_MultaJuros = ""
    
    'Verificar em objTituloReceber se há valor > 0 em PercMulta
    If objTituloReceber.dPercMulta > 0 Then
      
        'Se houver -> Montar instrução com valor da Multa
        sInstrucao_MultaJuros = "Após o vencimento multa de " & sSimbolo & " " & Format(objParcelaReceber.dSaldo * objTituloReceber.dPercMulta, "Standard")
        
    End If
    
    'Verificar em objTituloReceber se há valor > 0 em PercJurosDiarios
    If objTituloReceber.dPercJurosDiario > 0 Then
      
        'Se houver -> Montar instrução com valor do Juros
        If Len(sInstrucao_MultaJuros) = 0 Then
            
            sInstrucao_MultaJuros = sInstrucao_MultaJuros & "Após o vencimento "
            
        Else
            
            sInstrucao_MultaJuros = sInstrucao_MultaJuros & " e "
            
        End If
        
        sInstrucao_MultaJuros = sInstrucao_MultaJuros & "juros de " & sSimbolo & " " & Format((objParcelaReceber.dSaldo * objTituloReceber.dPercJurosDiario), "Standard") & " ao dia"
       
    End If
    
    If Len(Trim(sInstrucao_MultaJuros)) <> 0 Then
    
        sInstrucoes = sInstrucoes & ". ATENÇÃO: " & sInstrucao_MultaJuros
    
    End If
    
    Preenche_Instrucoes_Boleto = SUCESSO
    
    Exit Function
    
Erro_Preenche_Instrucoes_Boleto:

    Preenche_Instrucoes_Boleto = gErr

    Select Case gErr

        Case 140884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 140885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESCRITEMNF_BOLETO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184419)

    End Select
    
    'Fecha Comando
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

