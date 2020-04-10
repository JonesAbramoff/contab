VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Gera_Sped_Contabil_Diario(ByVal sDiretorio As String, ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lNumOrd As Long, ByVal objTela As Object, ByVal sContaOutros As String, ByVal iTipo As Integer, ByVal iIndSitPeriodo As Integer, ByVal iIndSitEspecial As Integer, Optional ByVal iCodVersao As Integer = 1, Optional ByVal iIndNIRE As Integer = 0, Optional ByVal iFinalidade As Integer = 0, Optional ByVal sHashEscrSubst As String = "", Optional ByVal sNIRESubst As String = "", Optional ByVal iEmpGrandePorte As Integer = 0, Optional ByVal iTipoECD As Integer = 0, Optional ByVal sCodSCP As String = "", Optional ByVal colSCPs As Collection = Nothing) As Long

Dim iIndice As Integer, lErro As Long, lNumIntSped As Long, objECD As ClassSpedContabil, lSeq As Long, objFilialEmpresa As New AdmFiliais
Dim objEndereco As New ClassEndereco, objCidades As New ClassCidades, lTransacao As Long, alComando(1 To 41) As Long, sCodigoInst As String
Dim sCodigoCad As String, tPlanoConta As typePlanoConta, iNaturezaSped As Integer, sContaTela As String, sContaPai As String, lModelo As Long
Dim sContaRef As String, sCclIni As String, sCclFim As String, sContaTelaRef As String, sCclIniDem As String, sCclFimDem As String, iCodigoDem As Integer
Dim sCcl As String, sCclMascarado As String, iUsoCcl As Integer, sDescCcl As String, dSaldoIni As Double, dSaldoDeb As Double, dSaldoCre As Double, dSaldoFim As Double
Dim sConta As String, tLanc As typeLancamento, sOrigemAtual As String, iExercicioAtual As Integer, iPeriodoLanAtual As Integer, lDocAtual As Long
Dim iFilialEmpresaAtual As Integer, sIdentificador As String, lTotalRegBloco1 As Long
Dim sDescQualiSigCTB As String, objReg As ClassSpedRegistro, colSpedRegistro As New Collection
Dim objRegI050 As ClassSpedRegistro, objRegI051 As ClassSpedRegistro, objRegI052 As ClassSpedRegistro
Dim objRegI200 As ClassSpedRegistro, objRegI250 As ClassSpedRegistro
Dim sRegistro As String, iPos As Integer, dtData As Date, sCclMascarada As String, iExerAnt As Integer, objPer As New ClassPeriodo
Dim sContaAtual As String, dSaldoAcum As Double, dSaldoCreAcum As Double, dSaldoDebAcum As Double, vbMsgRes As VbMsgBoxResult
Dim lContador As Long, dSaldoIniTotal As Double, dSaldoFimTotal As Double, sContaTelaUltimo As String, dSaldoIniUltimo As Double, dSaldoFimUltimo As Double
Dim lSeqUltimo As Long, dSaldoCreUltimo As Double, dSaldoDebUltimo As Double, iX1 As Double, iX2 As Double, iX3 As Double
Dim dtDataAtual As Date, dValorDocPos As Double, dValorDocNeg As Double, objLanc As ClassLancamento_Detalhe, colLanc As New Collection
Dim colSpedContabil1 As New Collection, objECD1 As ClassSpedContabil1, objECD1_Aux As ClassSpedContabil1, objECDAux1 As ClassSpedContabil
Dim dValorDoc As Double, dTotalSldIniCred As Double, dTotalSldIniDeb As Double, dTotalSldFimCred As Double, dTotalSldFimDeb As Double
Dim iContador As Integer, dtDataIniAux As Date, dtDataFimAux As Date, tRelDRE As typeRelDRE, dValor As Double, colSpedCont1 As New Collection
Dim objSpedCont1 As ClassSpedContabil1, iAchou As Integer, lTotalRegBlocoJ As Long, lTotalRegBlocoM As Long
Dim objRegI150 As ClassSpedRegistro, objRegI155 As ClassSpedRegistro, objECD1_Prox As ClassSpedContabil1
Dim colSpedContRes1 As New Collection, objSpedContRes1 As ClassSpedContabil1, colSpedContRes2 As New Collection, objSpedContRes2 As ClassSpedContabil1
Dim iAchou2 As Integer, lTotal As Long, sContaEnxuta As String, iAchou1 As Integer, dDiferenca As Double
Dim colSpedContabil1_Lanc As New Collection, objECD1_Lanc As ClassSpedContabil1
Dim lSeq155Final As Long, objECD1_Outros As ClassSpedContabil1, colM025 As New Collection, vRegM025 As Variant, colM025Aux As New Collection, vRegM025Aux As Variant
Dim sContaAnt As String, bSemCcl As Boolean, dLucro As Double, iAtivoCcl As Integer, iAtivoCta As Integer, bTrata As Boolean, colCclInatSemMov As New Collection
Dim colCtaInatSemMov As New Collection, vsCcl As Variant, vsCta As Variant, lAux As Long
Dim colCclCta As Collection, colCtaSemCcl As New Collection, sCclAux As String, sIE As String, sIM As String
Dim dSaldoIniCta As Double, dSaldoIniCtaCcl As Double, objPerDataIni As New ClassPeriodo, dValorIni As Double
Dim objSCP As AdmFiliais, sLivro As String, iTipoModelo As Integer
Dim objExercicio As New ClassExercicio, bIgnoraTesteCcl As Boolean, dValorExercAnt As Double

On Error GoTo Erro_Gera_Sped_Contabil_Diario

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 203113

    'desativa os locks dos comandos a seguir
    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 207913

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203114
    Next

    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_SPEDCONTABIL", lNumIntSped)
    If lErro <> SUCESSO Then gError 203115
    
    lErro = Comando_Executar(alComando(9), "SELECT UsoCcl From Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 203173
    
    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203174
    
    '*********** totalizador de registros *************************
    Call objTela.SPED_Status("Totalizando Registros", "Plano Conta")
    
    lErro = Comando_Executar(alComando(21), "SELECT Count(*) From PlanoConta", lContador)
    If lErro <> AD_SQL_SUCESSO Then gError 203254
    
    lErro = Comando_BuscarPrimeiro(alComando(21))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203255
    
    objTela.dValorTotal = lContador
    
    lErro = CF("Periodo_Le", dtDataIni, objPerDataIni)
    If lErro <> SUCESSO Then gError 203059
    
    lErro = CF("Exercicio_Le_Data", dtDataIni, objExercicio)
    If lErro <> SUCESSO Then gError 203059
    
    If iUsoCcl = 2 Then
    
        Call objTela.SPED_Status("Totalizando Registros", "Ccl")
    
        lErro = Comando_Executar(alComando(21), "SELECT Count(*) FROM Ccl", lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 203256
        
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203257
        
        objTela.dValorTotal = objTela.dValorTotal + lContador
    
        Call objTela.SPED_Status("Totalizando Registros", "Conta\Ccl")
    
        lErro = Comando_Executar(alComando(21), "SELECT COUNT(*) FROM (SELECT Conta, Ccl From MvPerCcl GROUP BY Conta, Ccl) As A ", lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 203258
        
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203259
        
        objTela.dValorTotal = objTela.dValorTotal + lContador
    
    Else
    
        Call objTela.SPED_Status("Totalizando Registros", "Conta")
    
        lErro = Comando_Executar(alComando(21), "SELECT Count(*) From PlanoConta WHERE TipoConta = 3", lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 203260
        
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203261
        
        objTela.dValorTotal = objTela.dValorTotal + lContador
    
    End If
        
    Call objTela.SPED_Status("Totalizando Registros", "Lançamentos")

    If iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(21), "SELECT  Count(*) FROM Lancamentos WHERE FilialEmpresa < ? AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) ", lContador, 50, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 207327
    Else
        lErro = Comando_Executar(alComando(21), "SELECT  Count(*) FROM Lancamentos WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) ", lContador, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203262
    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(21))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203263
    
    objTela.dValorTotal = objTela.dValorTotal + lContador

    '*********** inicio *****************
    Call objTela.SPED_Status("Processando", "Modelo Referencial")
    
    Set objECD = New ClassSpedContabil
    
    If iTipo = SPED_CONTAB_TIPO_NORMAL Then
        
        ' descobre o plano de contas referencial que sera usado
        lErro = Comando_Executar(alComando(5), "SELECT Codigo, Tipo FROM PlanoContaRefModelo WHERE Oficial = 1 AND AnoVigencia <= ? ORDER BY AnoVigencia DESC", _
        lModelo, iTipoModelo, Year(dtDataFim))
        If lErro <> AD_SQL_SUCESSO Then gError 203116
    
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203117
        
        If lErro <> AD_SQL_SUCESSO Then gError 203118
        
    Else
        iTipoModelo = 10
        lModelo = 1
    End If
    
    '******** registro 0000 *****************
    Call objTela.SPED_Status("Processando Registro 0000", "")
    
    If iTipo = SPED_CONTAB_TIPO_NORMAL Then
        objECD.sRegistro = "|0000|LECD|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
    Else
        objECD.sRegistro = "|0000|LALU|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
    End If
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    If objFilialEmpresa.iCodFilial = EMPRESA_TODA Then objFilialEmpresa.iCodFilial = 1
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 203119

    If lErro <> SUCESSO Then gError 203120

    objECD.sRegistro = objECD.sRegistro & gsNomeEmpresa & "|" & objFilialEmpresa.sCgc & "|" & objFilialEmpresa.objEndereco.sSiglaEstado & "|"
    
    objCidades.sDescricao = objFilialEmpresa.objEndereco.sCidade
    
    lErro = CF("Cidade_Le_Nome", objCidades)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203121
    
    If lErro = ERRO_OBJETO_NAO_CADASTRADO Then gError 203122
    
    sIE = ""
    Call Formata_String_Numero(objFilialEmpresa.sInscricaoEstadual, sIE)
    
    sIM = ""
    Call Formata_String_Numero(objFilialEmpresa.sInscricaoMunicipal, sIM)
    
    objECD.sRegistro = objECD.sRegistro & sIE & "|" & objCidades.sCodIBGE & "|" & sIM & "|" & IIf(iIndSitEspecial <> 0, CStr(iIndSitEspecial), "") & "|"
    
    If iTipo = SPED_CONTAB_TIPO_FCONT Then
        objECD.sRegistro = objECD.sRegistro & CStr(iIndSitPeriodo) & "|"
    ElseIf iCodVersao >= 2 And iCodVersao < 5 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        objECD.sRegistro = objECD.sRegistro & CStr(iIndSitPeriodo) & "|" & CStr(iIndNIRE) & "|" & CStr(iFinalidade) & "|" & CStr(sHashEscrSubst) & "|" & CStr(sNIRESubst) & "|" & CStr(iEmpGrandePorte) & "|"
    ElseIf iCodVersao >= 5 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        objECD.sRegistro = objECD.sRegistro & CStr(iIndSitPeriodo) & "|" & CStr(iIndNIRE) & "|" & CStr(iFinalidade) & "|" & CStr(sHashEscrSubst) & "|" & CStr(iEmpGrandePorte) & "|"
    End If
    
    If iCodVersao >= 3 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        'TIP_ECD Indicador do tipo de
        'ECD:
        '0 – ECD de empresa não participante de SCP como sócio ostensivo.
        '1 – ECD de empresa participante de SCP como sócio ostensivo.
        '2 – ECD da SCP.
        objECD.sRegistro = objECD.sRegistro & CStr(iTipoECD) & "|"
    
        'COD_SCP Identificação da SCP
        '(CNPJ – art. 52 da Instrução Normativa RFB no 1.470, de 30 de maio de 2014).
        'Obs.: Só deve ser preenchido pela própria SCP com o CNPJ/Código da SCP (não é preenchido pelo sócio ostensivo).
        objECD.sRegistro = objECD.sRegistro & sCodSCP & "|"
    End If
        
    If iCodVersao >= 4 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        objECD.sRegistro = objECD.sRegistro & "N|"
    End If
    
    If iCodVersao >= 5 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        objECD.sRegistro = objECD.sRegistro & "N|"
    End If
        
    objECD.sTipo = "0000"
    objECD.iFilialEmpresa = iFilialEmpresa
    objECD.lNumIntSped = lNumIntSped
    objECD.lSeq = 1
    
    Set objReg = New ClassSpedRegistro
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    colSpedRegistro.Add objReg
    
    If Len(objECD.sRegistro) > 500 Then gError 203123
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203124

    If iTipo = SPED_CONTAB_TIPO_NORMAL Then
    
        '******** registro 0001 *****************
        Call objTela.SPED_Status("Processando Registro 0001", "")
    
        objECD.sTipo = "0001"
        objECD.lSeq = objECD.lSeq + 1
        
        objECD.sRegistro = "|0001|0|"
        
        Set objReg = New ClassSpedRegistro
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        colSpedRegistro.Add objReg
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203125
    
        '******** registro 0007 *****************
        Call objTela.SPED_Status("Processando Registro 0007", "")
    
        objECD.sTipo = "0007"
        
        Set objReg = New ClassSpedRegistro
        
        objReg.sTipo = objECD.sTipo
    
        colSpedRegistro.Add objReg
        
        sCodigoInst = String(STRING_SPED_CODIGOINST, 0)
        sCodigoCad = String(STRING_SPED_CODIGOCAD, 0)
        
        If iFilialEmpresa = EMPRESA_TODA Then
            lErro = Comando_Executar(alComando(2), "SELECT CodigoInst, CodigoCad FROM SpedCadastro WHERE FilialEmpresa = ?", _
            sCodigoInst, sCodigoCad, 1)
            If lErro <> AD_SQL_SUCESSO Then gError 203126
        Else
            lErro = Comando_Executar(alComando(2), "SELECT CodigoInst, CodigoCad FROM SpedCadastro WHERE FilialEmpresa = ?", _
            sCodigoInst, sCodigoCad, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 203126
        End If
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203127
    
        Do While lErro = AD_SQL_SUCESSO
        
            objECD.lSeq = objECD.lSeq + 1
        
            Call objTela.SPED_Status("Processando Registro 0007", CStr(objECD.lSeq) & SEPARADOR & sCodigoInst)
        
            objReg.lCount = objReg.lCount + 1
        
            objECD.sRegistro = "|0007|" & sCodigoInst & "|" & sCodigoCad & "|"
        
            If Len(objECD.sRegistro) > 500 Then gError 203128
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203129
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203130
            
        Loop
        
        '******** registro 0035 *****************
        If iCodVersao >= 3 And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        
            If colSCPs.Count > 0 Then
            
                Call objTela.SPED_Status("Processando Registro 0035", "")
            
                objECD.sTipo = "0035"
                
                Set objReg = New ClassSpedRegistro
                
                objReg.sTipo = objECD.sTipo
            
                colSpedRegistro.Add objReg
                      
                For Each objSCP In colSCPs
                
                    objECD.lSeq = objECD.lSeq + 1
                
                    Call objTela.SPED_Status("Processando Registro 0035", CStr(objECD.lSeq) & SEPARADOR & objSCP.sCgc)
                
                    objReg.lCount = objReg.lCount + 1
                
                    objECD.sRegistro = "|0035|" & objSCP.sCgc & "|" & objSCP.sNome & "|"
                
                    If Len(objECD.sRegistro) > 500 Then gError 203128
                
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                    If lErro <> SUCESSO Then gError 203129
                    
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203130
                    
                Next
                
            End If
            
        End If
    
        '******** registro 0990 *****************
        Call objTela.SPED_Status("Processando Registro 0990", "")
    
        objECD.sTipo = "0990"
        objECD.lSeq = objECD.lSeq + 1
        
        Set objReg = New ClassSpedRegistro
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        colSpedRegistro.Add objReg
        
        objECD.sRegistro = "|0990|" & CStr(objECD.lSeq) & "|"
        
        If Len(objECD.sRegistro) > 500 Then gError 203131
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203132

    End If

    '******** registro I001 *****************
    Call objTela.SPED_Status("Processando Registro I001", "")

    objECD.sTipo = "I001"
    objECD.lSeq = objECD.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objReg = New ClassSpedRegistro
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    colSpedRegistro.Add objReg
    
    objECD.sRegistro = "|I001|0|"
    
    If Len(objECD.sRegistro) > 500 Then gError 203133
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203134

    If iTipo = SPED_CONTAB_TIPO_NORMAL Then
    
        '******** registro I010 *****************
        Call objTela.SPED_Status("Processando Registro I010", "")
    
        objECD.sTipo = "I010"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        sLivro = "G"
        If iTipoECD = 2 Then
            sLivro = "S"
        End If
        
        Select Case iCodVersao
            Case Is >= 5
                objECD.sRegistro = "|I010|" & sLivro & "|" & CStr(iCodVersao) & ".00|"
            Case 4
                objECD.sRegistro = "|I010|" & sLivro & "|4.00|"
            Case 3
                objECD.sRegistro = "|I010|" & sLivro & "|3.00|"
            Case 2
                objECD.sRegistro = "|I010|" & sLivro & "|2.00|"
            Case Else
                objECD.sRegistro = "|I010|" & sLivro & "|1.00|"
        End Select
        
        If Len(objECD.sRegistro) > 500 Then gError 203135
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203136
    
        '******** registro I030 *****************
        Call objTela.SPED_Status("Processando Registro I030", "")
    
        objECD.sTipo = "I030"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|I030|TERMO DE ABERTURA|" & lNumOrd & "|Escrituração Contábil Digital do Livro Diário Geral||"
        objECD.sRegistro = objECD.sRegistro & gsNomeEmpresa & "|" & objFilialEmpresa.sJucerja & "|"
        objECD.sRegistro = objECD.sRegistro & objFilialEmpresa.sCgc & "|" & IIf(objFilialEmpresa.dtDataJucerja = DATA_NULA And iCodVersao >= 2 And iTipo = SPED_CONTAB_TIPO_NORMAL, "", Format(objFilialEmpresa.dtDataJucerja, "ddmmyyyy")) & "||"
        
        If iCodVersao = 2 Then
            objECD.sRegistro = objECD.sRegistro & "||||"
        End If
        If iCodVersao >= 3 Then
            objECD.sRegistro = objECD.sRegistro & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
        End If
        
        If Len(objECD.sRegistro) > 500 Then gError 203137
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203138

    End If
    
    '==============================================
    'Obtém ccls que não serão tratados
    sCcl = String(STRING_CCL, 0)
    sDescCcl = String(STRING_CCL_DESCRICAO, 0)

    lErro = Comando_Executar(alComando(10), "SELECT Ccl, DescCcl, AtivoCcl FROM Ccl ORDER BY Ccl", sCcl, sDescCcl, iAtivoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 203175

    lErro = Comando_BuscarPrimeiro(alComando(10))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203176

    Do While lErro = AD_SQL_SUCESSO
    
        bTrata = True
        If iAtivoCcl = DESMARCADO Then
            'Inativo
            lErro = Comando_Executar(alComando(36), "SELECT FilialEmpresa FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Data BETWEEN ? AND ? ", lAux, iFilialEmpresa, sCcl, objExercicio.dtDataInicio, objExercicio.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 203254
            
            lErro = Comando_BuscarPrimeiro(alComando(36))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203255
            
            If lErro = AD_SQL_SEM_DADOS Then
                'Sem movimento no período
                bTrata = False
                colCclInatSemMov.Add sCcl
            End If
        End If
        
        lErro = Comando_BuscarProximo(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203176
        
    Loop
    '==============================================

    '******** registro I050 *****************

    Call objTela.SPED_Status("Processando Registro I050", "")

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT Conta, TipoConta, NivelConta, DescConta, NaturezaSped, Ativo FROM PlanoConta ORDER BY Conta", _
    tPlanoConta.sConta, tPlanoConta.iTipoConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, iNaturezaSped, iAtivoCta)
    If lErro <> AD_SQL_SUCESSO Then gError 203139

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203140

    Set objRegI050 = New ClassSpedRegistro
    colSpedRegistro.Add objRegI050
    objRegI050.sTipo = "I050"
    
    Set objRegI051 = New ClassSpedRegistro
    colSpedRegistro.Add objRegI051
    objRegI051.sTipo = "I051"

    Set objRegI052 = New ClassSpedRegistro
    colSpedRegistro.Add objRegI052
    objRegI052.sTipo = "I052"

    Do While lErro = AD_SQL_SUCESSO
    
        bTrata = True
        If iAtivoCta <> CONTA_ATIVA Then
            'Inativo
            'Verifica se tem lançamentos para própria ou filhos, se tiver não pode deixar de colocar no arquivo
            lErro = Comando_Executar(alComando(36), "SELECT FilialEmpresa FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta LIKE ? AND Data BETWEEN ? AND ? ", lAux, iFilialEmpresa, tPlanoConta.sConta & "%", objExercicio.dtDataInicio, objExercicio.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 203254
            
            lErro = Comando_BuscarPrimeiro(alComando(36))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203255
            
            If lErro = AD_SQL_SEM_DADOS Then
                'Sem movimento no período
                bTrata = False
                colCtaInatSemMov.Add tPlanoConta.sConta
            End If
            
        End If
        
        If bTrata Then
            
            Set colSpedCont1 = New Collection
        
            Call objTela.SPED_Status("Processando Registro I050", CStr(objECD.lSeq + 1) & SEPARADOR & tPlanoConta.sConta)
        
    '        If tPlanoConta.iNivelConta = 1 Then iNatSpedAtual = iNaturezaSped
        
            If iTipo = SPED_CONTAB_TIPO_NORMAL Then
                lErro = Comando_Executar(alComando(4), "SELECT DataAtualizacao FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                    dtData, tPlanoConta.sConta, dtDataFim)
            Else
                lErro = Comando_Executar(alComando(4), "SELECT DataAtualizacao FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                    dtData, tPlanoConta.sConta, dtDataIni)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 203141
        
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203142
        
            If lErro = AD_SQL_SEM_DADOS Then dtData = StrParaDate("01/01/2009")
        
            objECD.sTipo = "I050"
            objECD.lSeq = objECD.lSeq + 1
            lTotalRegBloco1 = lTotalRegBloco1 + 1
        
            objRegI050.lCount = objRegI050.lCount + 1
        
            objECD.sRegistro = "|I050|" & Format(dtData, "ddmmyyyy") & "|" & Format(iNaturezaSped, "00") & "|" & IIf(tPlanoConta.iTipoConta = 1, "S", "A") & "|"
            objECD.sRegistro = objECD.sRegistro & tPlanoConta.iNivelConta & "|"
            
            lErro = Mascara_MascararConta(tPlanoConta.sConta, sContaTela)
            If lErro <> SUCESSO Then gError 203143
            
            If tPlanoConta.iNivelConta > 1 Then
                
                lErro = Mascara_RetornaContaPai(tPlanoConta.sConta, sConta)
                If lErro <> SUCESSO Then gError 203144
            
                lErro = Mascara_MascararConta(sConta, sContaPai)
                If lErro <> SUCESSO Then gError 203143
            
            Else
                sContaPai = ""
            End If
            
            objECD.sRegistro = objECD.sRegistro & sContaTela & "|" & sContaPai & "|" & Trim(tPlanoConta.sDescConta) & "|"
        
            If Len(objECD.sRegistro) > 500 Then gError 203145
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203146
            
            'se for conta analitica e de resultado
            'guarda as contas de resultado
            If tPlanoConta.iTipoConta = 3 And iNaturezaSped = 4 Then
                Set objSpedContRes1 = New ClassSpedContabil1
                
                objSpedContRes1.sConta = tPlanoConta.sConta
                
                colSpedContRes1.Add objSpedContRes1
            End If
            
            objTela.dValorAtual = objTela.dValorAtual + 1
            objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
        
            If objTela.dValorAtual < objTela.dValorTotal Then objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
        
            DoEvents
            
            If objTela.iCancelaBatch = CANCELA_BATCH Then
        
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
            
                If vbMsgRes = vbYes Then gError 203265
                
                objTela.iCancelaBatch = 0
                    
            End If
            
            'se for uma conta analitica
            If tPlanoConta.iTipoConta = 3 Then
            
                Set colCclCta = New Collection
            
                Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objECD.lSeq) & SEPARADOR & tPlanoConta.sConta)
                
                bIgnoraTesteCcl = False
                If UCase(left(gsNomeEmpresa, 5)) = "INPAL" And (left(tPlanoConta.sConta, 1) = "1" Or left(tPlanoConta.sConta, 1) = "2") Then
                    bSemCcl = True
                    bIgnoraTesteCcl = True
                    colCtaSemCcl.Add tPlanoConta.sConta
                End If
                
                If Not bIgnoraTesteCcl Then
                    
                    sCcl = String(STRING_CCL, 0)
                    
                    lErro = Comando_Executar(alComando(34), "SELECT Ccl FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl <> ? AND Conta = ? AND Data BETWEEN ? AND ?  ", _
                    sCcl, iFilialEmpresa, "", tPlanoConta.sConta, objExercicio.dtDataInicio, objExercicio.dtDataFim)
                    If lErro <> AD_SQL_SUCESSO Then gError 203149
                    
                    lErro = Comando_BuscarPrimeiro(alComando(34))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                    
                    If lErro = AD_SQL_SUCESSO Then
                    
                        bSemCcl = False ' Essa conta trabalha com centro de custo
                        'If iTipo = SPED_CONTAB_TIPO_FCONT Then
                        'O ECD para o ECF tem o mesmo critério do FCONT, ou seja, ou estpa perfeito por ccl ou não pode ter ccl
                        
                            sCcl = String(STRING_CCL, 0)
                        
                            'Se achou lançamento com ccl procura sem ccl para mesma conta/período
                            lErro = Comando_Executar(alComando(38), "SELECT Ccl FROM Lancamentos WHERE Ccl = ? AND Conta = ? AND (FilialEmpresa = ? OR ?=?) AND Data BETWEEN ? AND ? AND ABS(Valor) > 0.001  AND (Aglutinado = 0 OR Aglutinado = 1) AND Origem NOT IN ('APE','APP','EAE','EAP') ", _
                            sCcl, "", tPlanoConta.sConta, iFilialEmpresa, iFilialEmpresa, EMPRESA_TODA, objExercicio.dtDataInicio, objExercicio.dtDataFim)
                            If lErro <> AD_SQL_SUCESSO Then gError 203149
                    
                            lErro = Comando_BuscarPrimeiro(alComando(38))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
        
                            If lErro = AD_SQL_SUCESSO Then
                                bSemCcl = True 'Não trabalha integralmente
                                colCtaSemCcl.Add tPlanoConta.sConta
                            Else
                                'Testa o saldo inicial com CCL e total, se não forem iguais não pode considerar que a conta trabalha com ccl
                                lErro = Comando_Executar(alComando(40), "SELECT SldIni FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? ", _
                                dSaldoIniCta, iFilialEmpresa, objPerDataIni.iExercicio, tPlanoConta.sConta)
                                If lErro <> AD_SQL_SUCESSO Then gError 203149
                        
                                lErro = Comando_BuscarPrimeiro(alComando(40))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                                
                                lErro = Comando_Executar(alComando(41), "SELECT SUM(SldIni) FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? ", _
                                dSaldoIniCtaCcl, iFilialEmpresa, objPerDataIni.iExercicio, tPlanoConta.sConta)
                                If lErro <> AD_SQL_SUCESSO Then gError 203149
                        
                                lErro = Comando_BuscarPrimeiro(alComando(41))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                                
                                If Abs(dSaldoIniCta - dSaldoIniCtaCcl) > DELTA_VALORMONETARIO Then
                                    bSemCcl = True 'Não trabalha integralmente
                                    colCtaSemCcl.Add tPlanoConta.sConta
                                End If
                            
                            End If
                            
                        'End If
                        
                    Else
                        bSemCcl = True 'Não existe movimentação com ccl no período
                        colCtaSemCcl.Add tPlanoConta.sConta
                    End If
                    
                End If
                
                'Registro I051
                'If iCodVersao < 2 Or iTipo <> SPED_CONTAB_TIPO_NORMAL Then
                
                sContaRef = String(STRING_CONTA, 0)
                sContaTelaRef = String(STRING_CONTA, 0)
                sCclIni = String(STRING_CCL, 0)
                sCclFim = String(STRING_CCL, 0)
            
                lErro = Comando_Executar(alComando(6), "SELECT P.ContaImp, PC.Conta, PC.CclInicial, PC.CclFinal FROM PlanoContaRefConfig AS PC, PlanoContaRef As P WHERE PC.CodigoModelo = P.CodigoModelo AND PC.Conta = P.Conta AND PC.CodigoModelo = ? AND ContaInicial <= ? And ContaFinal >= ? ", _
                sContaTelaRef, sContaRef, sCclIni, sCclFim, lModelo, tPlanoConta.sConta, tPlanoConta.sConta)
                If lErro <> AD_SQL_SUCESSO Then gError 203147
            
                lErro = Comando_BuscarPrimeiro(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203148
            
                iX1 = 0
                Do While lErro = AD_SQL_SUCESSO
                    
                    iX1 = iX1 + 1
                
                    Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objECD.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sContaTelaRef)
                               
                    'Inclui a associação por ccl se o ccl estiver associado a conta
                    If Not bSemCcl Then
                        'se o centro de custo da conta  estiver preenchido
                    
                        sCcl = String(STRING_CCL, 0)
                        
                        lErro = Comando_Executar(alComando(33), "SELECT DISTINCT Ccl FROM MvPerCcl WHERE FilialEmpresa = ? AND (Ccl >= ? OR ? = ?) AND (Ccl <= ? OR ? = ?) AND Conta = ? AND Exercicio = ? ORDER BY Ccl", _
                        sCcl, iFilialEmpresa, sCclIni, sCclIni, "", sCclFim, sCclFim, "", tPlanoConta.sConta, objExercicio.iExercicio)
                        If lErro <> AD_SQL_SUCESSO Then gError 203149
                        
                        lErro = Comando_BuscarPrimeiro(alComando(33))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                        
                        iX2 = 0
                        Do While lErro = AD_SQL_SUCESSO
                        
                            bTrata = True
                            For Each vsCcl In colCclInatSemMov
                                If vsCcl = sCcl Then
                                    bTrata = False 'O ccl estava inativo e sem movimento no período, não pode colocar ele no arquivo senão vai dar erro
                                    Exit For
                                End If
                            Next
                            For Each vsCcl In colCclCta
                                If vsCcl = sCcl Then
                                    bTrata = False 'Se já tratou esse ccl, não duplica
                                    Exit For
                                End If
                            Next
                            
                            If bTrata Then
                                                           
                                iX2 = iX2 + 1
                                 
                                If colCclCta.Count = 0 Then
                                    'Adiciona uma referência temporária para poder colocar o saldo sem ccl
                                    vRegM025Aux = "|M025|" & sContaTela & "||" & sContaTelaRef & "|"
                                    colM025Aux.Add vRegM025Aux
                                End If
                                 
                                colCclCta.Add sCcl
                                 
                                Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objECD.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sContaTelaRef & SEPARADOR & sCcl)
                                 
                                lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                                If lErro <> SUCESSO Then gError 203151
                    
                                vRegM025Aux = "|M025|" & sContaTela & "|" & sCclMascarado & "|" & sContaTelaRef & "|"
                                colM025Aux.Add vRegM025Aux
                         
                                objECD.lSeq = objECD.lSeq + 1
                                    
'                                If iCodVersao < 3 Then
'                                   objECD.sRegistro = "|I051|10|" & sCclMascarado & "|" & sContaTelaRef & "|"
'                                Else
'
'                                   'Campo 02 – Código da Instituição Responsável pela Manutenção do Plano de Contas Referencial
'                                   'Código da Instituição Responsável pela Manutenção do Plano de Contas Referencial
'                                   'Código Grupo / Conta
'                                   '1 PJ em Geral (L100A + L300A da ECF)
'                                   '2 PJ em Geral – Lucro Presumido (P100 + P150 da ECF)
'                                   '3 Financeiras (L100B + L300B da ECF)
'                                   '4 Seguradoras (L100C + L300C da ECF)
'                                   '5 Imunes e Isentas em Geral (U100A + U150A da ECF)
'                                   '6 Financeiras – Imunes e Isentas (U100B + U150B da ECF)
'                                   '7 Seguradoras – Imunes e Isentas (U100C + U150C da ECF)
'                                   '8 Entidades Fechadas de Previdência Complementar (U100D + U150D da ECF)
'                                   '9 Partidos Políticos (U100E + U150E da ECF)
'
'                                   If objFilialEmpresa.iLucroPresumido = MARCADO Then
'                                       objECD.sRegistro = "|I051|2|" & sCclMascarado & "|" & sContaTelaRef & "|"
'                                   Else
'                                       objECD.sRegistro = "|I051|1|" & sCclMascarado & "|" & sContaTelaRef & "|"
'                                   End If
'
'                                End If

                                objECD.sRegistro = "|I051|" & CStr(iTipoModelo) & "|" & sCclMascarado & "|" & sContaTelaRef & "|"
                                     
                                lTotalRegBloco1 = lTotalRegBloco1 + 1
                                 
                                objRegI051.lCount = objRegI051.lCount + 1
                             
                                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                                objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                                If lErro <> SUCESSO Then gError 203153
                                    
                            End If
                                
                            lErro = Comando_BuscarProximo(alComando(33))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203158
                            
                        Loop
                        
                    ElseIf bSemCcl And sContaAnt <> sContaTela Then ' Se a conta não tem ccl e ainda não está associad a uma conta referencial
                        
                        sContaAnt = sContaTela
                        
                        'se o centro de custo da conta estiver zerado
                        
                        vRegM025Aux = "|M025|" & sContaTela & "||" & sContaTelaRef & "|"
                        colM025Aux.Add vRegM025Aux
                        
                        'If iCodVersao < 2 Or iTipo <> SPED_CONTAB_TIPO_NORMAL Then
                            
                        objECD.sRegistro = "|I051|" & CStr(iTipoModelo) & "||" & sContaTelaRef & "|"
                           
                        If Len(objECD.sRegistro) > 500 Then gError 203160
                            
                                '******** registro I051 *****************
                        objECD.sTipo = "I051"
                        objECD.lSeq = objECD.lSeq + 1
                        lTotalRegBloco1 = lTotalRegBloco1 + 1
                        
                        objRegI051.lCount = objRegI051.lCount + 1
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                        If lErro <> SUCESSO Then gError 203161
                                
                        'End If
                    
                    End If
                    
                    lErro = Comando_BuscarProximo(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203148
                    
                Loop
'
'                End If
                   
                'TRATA O I052 (não pode ficar dentro do tratamento do I051 porque é independente, deve existir mesmo sem associação ao plano referencial
                'Inclui a associação por ccl se o ccl estiver associado a conta
                If Not bSemCcl Then
                    'se o centro de custo da conta  estiver preenchido
                
                    sCcl = String(STRING_CCL, 0)
                    
                    lErro = Comando_Executar(alComando(33), "SELECT DISTINCT Ccl FROM MvPerCcl WHERE FilialEmpresa = ? AND Conta = ? AND Exercicio = ? ORDER BY Ccl", _
                    sCcl, iFilialEmpresa, tPlanoConta.sConta, objExercicio.iExercicio)
                    If lErro <> AD_SQL_SUCESSO Then gError 203149
                                        
                    lErro = Comando_BuscarPrimeiro(alComando(33))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                    
                    iX2 = 0
                    Do While lErro = AD_SQL_SUCESSO
                    
                        bTrata = True
                        For Each vsCcl In colCclInatSemMov
                            If vsCcl = sCcl Then
                                bTrata = False 'O ccl estava inativo e sem movimento no período, não pode colocar ele no arquivo senão vai dar erro
                                Exit For
                            End If
                        Next
'                        For Each vsCcl In colCclCta
'                            If vsCcl = sCcl Then
'                                bTrata = False 'Se já tratou esse ccl, não duplica
'                                Exit For
'                            End If
'                        Next
                        
                        If bTrata Then
                        
                            iX2 = iX2 + 1
                            
                            colCclCta.Add sCcl
                            
                            Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objECD.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sContaTelaRef & SEPARADOR & sCcl)
                            
                            lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                            If lErro <> SUCESSO Then gError 203151
                                                    
                            'se for conta de ativo ou passivo a conta aglutinadora é a propria conta me questao
                            If (iNaturezaSped = 1 Or iNaturezaSped = 2 Or iNaturezaSped = 3) And tPlanoConta.iTipoConta = 3 Then
                                
                                iAchou = 0
                                For Each objSpedCont1 In colSpedCont1
                                    If objSpedCont1.sConta = sContaTela And objSpedCont1.sCcl = "" Then
                                        iAchou = 1
                                        Exit For
                                    End If
                                Next
                                
                                If iAchou = 0 Then
                                
                                    Set objSpedCont1 = New ClassSpedContabil1
                                    
                                    objSpedCont1.sConta = sContaTela
                                    objSpedCont1.sContaBD = tPlanoConta.sConta
                                    
                                    colSpedCont1.Add objSpedCont1
                                                            
                                End If
                                
                            Else
                        
                                'se for conta de demonstrativo ==> pegar o demonstrativo SPED para a conta em questao
                                lErro = Comando_Executar(alComando(7), "SELECT R.Codigo FROM RelDRE As R, RelDREConta As RC WHERE R.Relatorio = RC.Relatorio AND R.Modelo = RC.Modelo AND R.Codigo = RC.Codigo AND Tipo = 0 AND Upper(RC.Relatorio) = 'DRE' AND Upper(RC.Modelo) = 'SPED' AND ContaInicial <= ? And ContaFinal >= ? AND CclInicial <= ? AND CclFinal >=? ", _
                                iCodigoDem, tPlanoConta.sConta, tPlanoConta.sConta, sCcl, sCcl)
                                If lErro <> AD_SQL_SUCESSO Then gError 203155

                                lErro = Comando_BuscarPrimeiro(alComando(7))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203156

                                If lErro = AD_SQL_SUCESSO Then
                                    
                                    iAchou = 0
                                    For Each objSpedCont1 In colSpedCont1
                                        If objSpedCont1.sConta = iCodigoDem And objSpedCont1.sCcl = sCclMascarado Then
                                            iAchou = 1
                                            Exit For
                                        End If
                                    Next

                                    If iAchou = 0 Then

                                        Set objSpedCont1 = New ClassSpedContabil1

                                        objSpedCont1.sCcl = sCclMascarado
                                        objSpedCont1.sConta = iCodigoDem
                                        objSpedCont1.sContaBD = tPlanoConta.sConta

                                        colSpedCont1.Add objSpedCont1

                                        'lErro = SpedContabil_FormulaDRE(objECD, iCodigoDem, lTotalRegBloco1, objRegI052, colSpedCont1)
                                        'If lErro <> SUCESSO Then gError 207350

                                        lErro = Comando_BuscarProximo(alComando(7))
                                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207351

                                    End If
                            
                                Else
                                
                                    'procura no dre sem o ccl
                                    'ccl não configurado -> Vale para todos
                                    lErro = Comando_Executar(alComando(39), "SELECT R.Codigo FROM RelDRE As R, RelDREConta As RC WHERE R.Relatorio = RC.Relatorio AND R.Modelo = RC.Modelo AND R.Codigo = RC.Codigo AND Tipo = 0 AND Upper(RC.Relatorio) = 'DRE' AND Upper(RC.Modelo) = 'SPED' AND ContaInicial <= ? And ContaFinal >= ? AND CclInicial = '' AND CclFinal = ''", _
                                        iCodigoDem, tPlanoConta.sConta, tPlanoConta.sConta)
                                    If lErro <> AD_SQL_SUCESSO Then gError 203155

                                    lErro = Comando_BuscarPrimeiro(alComando(39))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203156

                                    If lErro = AD_SQL_SUCESSO Then

                                        iAchou = 0
                                        For Each objSpedCont1 In colSpedCont1
                                            If objSpedCont1.sConta = iCodigoDem And objSpedCont1.sCcl = "" Then
                                                iAchou = 1
                                                Exit For
                                            End If
                                        Next

                                        If iAchou = 0 Then

                                            Set objSpedCont1 = New ClassSpedContabil1

                                            objSpedCont1.sCcl = ""
                                            objSpedCont1.sConta = iCodigoDem
                                            objSpedCont1.sContaBD = tPlanoConta.sConta

                                            colSpedCont1.Add objSpedCont1

                                        End If

                                    End If
                                
                                End If
                            
                            End If
                            
                        End If
                        
                        lErro = Comando_BuscarProximo(alComando(33))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203158
                    
                    Loop
                
                ElseIf bSemCcl Then 'And sContaAnt <> sContaTela Then ' Se a conta não tem ccl e ainda não está associad a uma conta referencial
                
                    sContaAnt = sContaTela
                                            
                    'se for conta de ativo ou passivo a conta aglutinadora é a propria conta me questao
                    If (iNaturezaSped = 1 Or iNaturezaSped = 2 Or iNaturezaSped = 3) And tPlanoConta.iTipoConta = 3 Then
                        iAchou = 0
                        For Each objSpedCont1 In colSpedCont1
                            If objSpedCont1.sConta = sContaTela And objSpedCont1.sCcl = "" Then
                                iAchou = 1
                                Exit For
                            End If
                        Next
                        If iAchou = 0 Then
                            Set objSpedCont1 = New ClassSpedContabil1
                            objSpedCont1.sConta = sContaTela
                            objSpedCont1.sContaBD = tPlanoConta.sConta
                            colSpedCont1.Add objSpedCont1
                        End If
                    Else
                    
                        Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objECD.lSeq) & SEPARADOR & tPlanoConta.sConta)
                    
                        'se for conta de demonstrativo ==> pegar o demonstrativo SPED para a conta em questao
                        sCclIniDem = String(STRING_CCL, 0)
                        sCclFimDem = String(STRING_CCL, 0)
                    
                        lErro = Comando_Executar(alComando(7), "SELECT RC.Codigo, RC.CclInicial, RC.CclFinal FROM RelDRE As R, RelDREConta As RC WHERE R.Relatorio = RC.Relatorio AND R.Modelo = RC.Modelo AND R.Codigo = RC.Codigo AND Tipo = 0 AND Upper(RC.Relatorio) = 'DRE' AND Upper(RC.Modelo) = 'SPED' AND ContaInicial <= ? And ContaFinal >= ? ", _
                        iCodigoDem, sCclIniDem, sCclFimDem, tPlanoConta.sConta, tPlanoConta.sConta)
                        If lErro <> AD_SQL_SUCESSO Then gError 203163
                
                        lErro = Comando_BuscarPrimeiro(alComando(7))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203164
                        
                        iX1 = 0
                        Do While lErro = AD_SQL_SUCESSO
                        
                            iX1 = iX1 + 1
                    
                            Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objECD.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & CStr(iCodigoDem))
                    
                            If sCclIniDem = "" And sCclFimDem = "" Then
                                
'                                iAchou = 0
'                                For Each objSpedCont1 In colSpedCont1
'                                    If objSpedCont1.sConta = iCodigoDem And objSpedCont1.sCcl = "" Then
'                                        iAchou = 1
'                                        Exit For
'                                    End If
'                                Next
'
'                                If iAchou = 0 Then
'                                    Set objSpedCont1 = New ClassSpedContabil1
'                                    objSpedCont1.sConta = iCodigoDem
'                                    objSpedCont1.sContaBD = tPlanoConta.sConta
'                                    colSpedCont1.Add objSpedCont1
'                                End If
                                
                            Else
                            
                                If sCclIniDem <> "" And sCclFimDem <> "" Then
                                
                                    sCcl = String(STRING_CCL, 0)
                                
                                    lErro = Comando_Executar(alComando(8), "SELECT Ccl FROM Ccl WHERE TipoCcl = 3 And Ccl >= ? AND Ccl <= ?", _
                                    sCcl, sCclIniDem, sCclFimDem)
                                    If lErro <> AD_SQL_SUCESSO Then gError 203166
                                
                                    lErro = Comando_BuscarPrimeiro(alComando(8))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203167
                                    
                                    iX2 = 0
                                    Do While lErro = AD_SQL_SUCESSO
                                    
                                        iX2 = iX2 + 1
                                        
                                        Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objECD.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & CStr(iCodigoDem) & SEPARADOR & sCcl)
                                       
                                        bTrata = True
                                        For Each vsCcl In colCclInatSemMov
                                            If vsCcl = sCcl Then
                                                bTrata = False 'O ccl estava inativo e sem movimento no período, não trata
                                                Exit For
                                            End If
                                        Next
                                       
                                        If bTrata Then
                                       
                                            lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                                            If lErro <> SUCESSO Then gError 203168
                                                        
                                            iAchou = 0
                                            For Each objSpedCont1 In colSpedCont1
                                                If objSpedCont1.sConta = iCodigoDem And objSpedCont1.sCcl = sCclMascarado Then
                                                    iAchou = 1
                                                    Exit For
                                                End If
                                            Next
                                            If iAchou = 0 Then
                                                Set objSpedCont1 = New ClassSpedContabil1
                                                objSpedCont1.sConta = iCodigoDem
                                                objSpedCont1.sContaBD = tPlanoConta.sConta
                                                objSpedCont1.sCcl = sCclMascarado
                                                colSpedCont1.Add objSpedCont1    '
                                            End If
                                    
                                        End If
                                    
                                        lErro = Comando_BuscarProximo(alComando(8))
                                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203170
                                            
                                    Loop
                                
                                End If
                            
                            End If
                            
                            'lErro = SpedContabil_FormulaDRE(objECD, iCodigoDem, lTotalRegBloco1, objRegI052, colSpedCont1)
                            'If lErro <> SUCESSO Then gError 207352
                            
                            lErro = Comando_BuscarProximo(alComando(7))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203171
                            
                        Loop
                    
                    End If
            
                End If
            
            End If
            

            If iTipo = SPED_CONTAB_TIPO_NORMAL Then
                For Each objSpedCont1 In colSpedCont1

                    objECD.sTipo = "I052"
                    objECD.lSeq = objECD.lSeq + 1
                    lTotalRegBloco1 = lTotalRegBloco1 + 1

                    objRegI052.lCount = objRegI052.lCount + 1
                    objECD.sRegistro = "|I052|" & objSpedCont1.sCcl & "|" & objSpedCont1.sConta & "|"

                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                    If lErro <> SUCESSO Then gError 203169

                Next

            End If
            
        End If
            
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203172
        
    Loop
    
    '******** registro I100 *****************
    Call objTela.SPED_Status("Processando Registro I100", "")

    'se usa ccusto extra-contabil
    If iUsoCcl = 2 Then
    
        sCcl = String(STRING_CCL, 0)
        sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    
        lErro = Comando_Executar(alComando(10), "SELECT Ccl, DescCcl, AtivoCcl FROM Ccl ORDER BY Ccl", sCcl, sDescCcl, iAtivoCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 203175
    
        lErro = Comando_BuscarPrimeiro(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203176
    
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
    
        objReg.sTipo = "I100"
    
        iX1 = 0
        Do While lErro = AD_SQL_SUCESSO
        
            bTrata = True
            For Each vsCcl In colCclInatSemMov
                If vsCcl = sCcl Then
                    bTrata = False 'O ccl estava inativo e sem movimento no período, não trata
                    Exit For
                End If
            Next
            
            If bTrata Then
            
                iX1 = iX1 + 1
            
                Call objTela.SPED_Status("Processando Registro I100 - Ccl", CStr(iX1) & SEPARADOR & sCcl)
            
                lErro = Comando_Executar(alComando(11), "SELECT DataAtualizacao FROM CclHistorico WHERE CodigoCcl = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                    dtData, sCcl, dtDataFim)
                If lErro <> AD_SQL_SUCESSO Then gError 203177
            
                lErro = Comando_BuscarPrimeiro(alComando(11))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203178
            
                If lErro <> AD_SQL_SUCESSO Then dtData = StrParaDate("01/01/2009")

                objECD.sTipo = "I100"
                objECD.lSeq = objECD.lSeq + 1
                lTotalRegBloco1 = lTotalRegBloco1 + 1
    
                objReg.lCount = objReg.lCount + 1
    
                objECD.sRegistro = "|I100|" & Format(dtData, "ddmmyyyy") & "|"
                
                lErro = Mascara_MascararCcl(sCcl, sCclMascarada)
                If lErro <> SUCESSO Then gError 203179
                
                objECD.sRegistro = objECD.sRegistro & sCclMascarada & "|" & sDescCcl & "|"
            
                If Len(objECD.sRegistro) > 500 Then gError 203180
            
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                If lErro <> SUCESSO Then gError 203181
    
                objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
                objTela.dValorAtual = objTela.dValorAtual + 1
            
                If objTela.dValorAtual < objTela.dValorTotal Then objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
            
                DoEvents
                
                If objTela.iCancelaBatch = CANCELA_BATCH Then
            
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
                
                    If vbMsgRes = vbYes Then gError 203266
                    
                    objTela.iCancelaBatch = 0
                        
                End If
                
            End If

            lErro = Comando_BuscarProximo(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203182
            
        Loop

    End If

    '******** registro I150 *****************
    dtDataIniAux = dtDataIni
        
    If (Month(dtDataFim) <> Month(dtDataIni) Or Year(dtDataFim) <> Year(dtDataIni)) And iTipo = SPED_CONTAB_TIPO_NORMAL Then
        dtDataFimAux = CDate("01/" & Month(dtDataIni) & "/" & Year(dtDataIni))
        dtDataFimAux = DateAdd("m", 1, dtDataFimAux)
        dtDataFimAux = DateAdd("d", -1, dtDataFimAux)
    Else
        dtDataFimAux = dtDataFim
    End If
    
    Set objRegI150 = New ClassSpedRegistro
    Set objRegI155 = New ClassSpedRegistro
    
    Do While dtDataIniAux >= dtDataIni And dtDataIniAux <= dtDataFim

        Call objTela.SPED_Status("Processando Registro I150", "")
        
        objECD.sTipo = "I150"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objECD.sRegistro = "|I150|" & Format(dtDataIniAux, "ddmmyyyy") & "|" & Format(dtDataFimAux, "ddmmyyyy") & "|"
    
        objRegI150.sTipo = objECD.sTipo
        objRegI150.lCount = objRegI150.lCount + 1
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203183
    
        '******** registro I155 *****************
        Call objTela.SPED_Status("Processando Registro I155", "")
    
        objECD.sTipo = "I155"
        
        tPlanoConta.sConta = String(STRING_CONTA, 0)
        
        objRegI155.sTipo = objECD.sTipo
        
        'se nao usa ccusto extra-contabil
        If iUsoCcl <> 2 Then
        
            lErro = CF("Periodo_Le", dtDataIniAux, objPer)
            If lErro <> SUCESSO Then gError 203059
        
            Call objTela.SPED_Status("Processando Registro I155", "Calcula os Saldos")
        
            'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
            lErro = CF("RelCtaSaldo_Prepara1", objPer, alComando(20), iExerAnt)
            If lErro <> SUCESSO Then gError 203060
        
            If iTipo = SPED_CONTAB_TIPO_FCONT Then
                lErro = Comando_Executar(alComando(12), "SELECT Conta, NaturezaSped FROM PlanoConta WHERE TipoConta = 3 AND NaturezaSped IN (1,2,3) ORDER BY Conta", _
                tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
            Else
                lErro = Comando_Executar(alComando(12), "SELECT Conta, NaturezaSped FROM PlanoConta WHERE TipoConta = 3 ORDER BY Conta", _
                tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 203184
        
            lErro = Comando_BuscarPrimeiro(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203185
        
            iX1 = 0
            Do While lErro = AD_SQL_SUCESSO
            
                bTrata = True
                For Each vsCta In colCtaInatSemMov
                    If vsCta = tPlanoConta.sConta Then
                        bTrata = False
                        Exit For
                    End If
                Next
                
                If tPlanoConta.sConta = "43010101000" Then
                    lErro = lErro
                End If
                
                If bTrata Then
                
                    iX1 = iX1 + 1
                
                    Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & tPlanoConta.sConta)
                
                    lErro = Conta_Saldo_Sped(iFilialEmpresa, tPlanoConta.sConta, dtDataIniAux, iExerAnt, objPer, dSaldoIni, alComando(27), alComando(28))
                    If lErro <> SUCESSO Then gError 203186
                    
                    lErro = Conta_Movimentacao_Sped(iFilialEmpresa, tPlanoConta.sConta, dtDataIniAux, dtDataFimAux, dSaldoCre, dSaldoDeb, alComando(29))
                    If lErro <> SUCESSO Then gError 203187
                    
                    If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
                
                        objECD.lSeq = objECD.lSeq + 1
                        lTotalRegBloco1 = lTotalRegBloco1 + 1
                        
                        lErro = Mascara_MascararConta(tPlanoConta.sConta, sContaTela)
                        If lErro <> SUCESSO Then gError 203188
                        
                        dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
                        
                        objECD.sRegistro = "|I155|" & sContaTela & "||" & Format(Abs(dSaldoIni), "fixed") & "|" & SFF_CD(dSaldoIni) & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & SFF_CD(dSaldoFim) & "|"
                        
                        objRegI155.lCount = objRegI155.lCount + 1
                        
                        lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, "", dSaldoIni, objECD, dtDataIniAux, dtDataFimAux, dtDataIni, tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
                        If lErro <> SUCESSO Then gError 207329
                
                        objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
                        objTela.dValorAtual = objTela.dValorAtual + 1
                    
                        If objTela.dValorAtual < objTela.dValorTotal Then objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
                    
                        DoEvents
                        
                        If objTela.iCancelaBatch = CANCELA_BATCH Then
                            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
                            If vbMsgRes = vbYes Then gError 203267
                            objTela.iCancelaBatch = 0
                        End If
                
                    End If
                    
                End If
            
                lErro = Comando_BuscarProximo(alComando(12))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203191
            
            Loop
            
        'se usa ccusto extra contabil
        Else
        
            'lErro = CF("Periodo_Le", dtDataIni, objPer)
            lErro = Periodo_Le_Sped(dtDataIniAux, objPer, alComando(24))
            If lErro <> SUCESSO Then gError 203068
            
            Call objTela.SPED_Status("Processando Registro I155", "Calcula os Saldos")
           
            'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
            lErro = CF("RelCtaSaldo_Prepara1", objPer, alComando(20), iExerAnt)
            If lErro <> SUCESSO Then gError 203069
        
            dSaldoIniTotal = 0
            dSaldoFimTotal = 0
        
            tPlanoConta.sConta = String(STRING_CONTA, 0)
        
            If iTipo = SPED_CONTAB_TIPO_FCONT Then
                lErro = Comando_Executar(alComando(22), "SELECT Conta, NaturezaSped FROM PlanoConta WHERE TipoConta = 3 AND NaturezaSped IN (1,2,3) ORDER BY Conta", _
                tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
            Else
                lErro = Comando_Executar(alComando(22), "SELECT Conta, NaturezaSped FROM PlanoConta WHERE TipoConta = 3 ORDER BY Conta", _
                tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 204745
        
            lErro = Comando_BuscarPrimeiro(alComando(22))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204746
        
            sContaAtual = tPlanoConta.sConta
            dSaldoAcum = 0
            dSaldoCreAcum = 0
            dSaldoDebAcum = 0
                
            iX1 = 0
            Do While lErro = AD_SQL_SUCESSO
                
                If tPlanoConta.sConta = "43010101000" Then
                    lErro = lErro
                End If
                
                bTrata = True
                For Each vsCta In colCtaInatSemMov
                    If vsCta = tPlanoConta.sConta Then
                        bTrata = False
                        Exit For
                    End If
                Next
                
                If bTrata Then
                     
                    iX1 = iX1 + 1
             
                    Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & tPlanoConta.sConta)
            
                    sConta = tPlanoConta.sConta

                    bSemCcl = False
                    For Each vsCta In colCtaSemCcl
                        If vsCta = tPlanoConta.sConta Then
                            bSemCcl = True
                            Exit For
                        End If
                    Next
                     
                    If Not bSemCcl Then
                    
                        sCcl = String(STRING_CCL, 0)
                 
                        lErro = Comando_Executar(alComando(13), "SELECT Ccl FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? GROUP BY Ccl ORDER BY Ccl", _
                        sCcl, iFilialEmpresa, objPer.iExercicio, tPlanoConta.sConta)
                         If lErro <> AD_SQL_SUCESSO Then gError 203192
                     
                         lErro = Comando_BuscarPrimeiro(alComando(13))
                         If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203193
                     
                         iX2 = 0
                         
                         Do While lErro = AD_SQL_SUCESSO
                         
                            bTrata = True
                            For Each vsCcl In colCclInatSemMov
                                If vsCcl = sCcl Then
                                    bTrata = False
                                    Exit For
                                End If
                            Next
                            
                            If bTrata Then
                            
                                iX2 = iX2 + 1
                        
                                Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sCcl)
                            
                                lErro = ContaCcl_Saldo_Sped(iFilialEmpresa, sConta, sCcl, dtDataIniAux, iExerAnt, objPer, dSaldoIni, alComando(25), alComando(26))
                                If lErro <> SUCESSO Then gError 203194
                                
                                lErro = ContaCcl_Movimentacao_Sped(iFilialEmpresa, sConta, sCcl, dtDataIniAux, dtDataFimAux, dSaldoCre, dSaldoDeb, alComando(27))
                                If lErro <> SUCESSO Then gError 203195
                                
                                If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
                            
                                    dSaldoCreAcum = Arredonda_Moeda(dSaldoCreAcum + dSaldoCre)
                                    dSaldoDebAcum = Arredonda_Moeda(dSaldoDebAcum + dSaldoDeb)
                                    dSaldoAcum = Arredonda_Moeda(dSaldoAcum + dSaldoIni)
                                    dSaldoIniTotal = dSaldoIniTotal + dSaldoIni
                                    
                                    objECD.lSeq = objECD.lSeq + 1
                                    lTotalRegBloco1 = lTotalRegBloco1 + 1
                        
                                    lErro = Mascara_MascararConta(sConta, sContaTela)
                                    If lErro <> SUCESSO Then gError 203196
                                    
                                    lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                                    If lErro <> SUCESSO Then gError 203197
                                    
                                    dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
                                    dSaldoFimTotal = dSaldoFimTotal + dSaldoFim
                                    
                                    objECD.sRegistro = "|I155|" & sContaTela & "|" & sCclMascarado & "|" & Format(Abs(dSaldoIni), "fixed") & "|" & SFF_CD(dSaldoIni) & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & SFF_CD(dSaldoFim) & "|"
                                
                                    lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, sCclMascarado, dSaldoIni, objECD, dtDataIniAux, dtDataFimAux, dtDataIni, tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
                                    If lErro <> SUCESSO Then gError 207328
                                    
                                    objRegI155.lCount = objRegI155.lCount + 1

                                End If
                         
                            End If
                         
                            objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
                            objTela.dValorAtual = objTela.dValorAtual + 1
                         
                            If objTela.dValorAtual < objTela.dValorTotal Then objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
                         
                            DoEvents
                             
                            If objTela.iCancelaBatch = CANCELA_BATCH Then
                                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
                                If vbMsgRes = vbYes Then gError 203268
                                objTela.iCancelaBatch = 0
                            End If
                         
                            lErro = Comando_BuscarProximo(alComando(13))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203200
                         
                         Loop
                         
                     End If
                     
                    lErro = Conta_Saldo_Sped(iFilialEmpresa, sContaAtual, dtDataIniAux, iExerAnt, objPer, dSaldoIni, alComando(27), alComando(28))
                    If lErro <> SUCESSO Then gError 203238
                    
                    lErro = Conta_Movimentacao_Sped(iFilialEmpresa, sContaAtual, dtDataIniAux, dtDataFimAux, dSaldoCre, dSaldoDeb, alComando(29))
                    If lErro <> SUCESSO Then gError 203239
                
                    dSaldoIni = Arredonda_Moeda(dSaldoIni - dSaldoAcum)
                    dSaldoCre = Arredonda_Moeda(dSaldoCre - dSaldoCreAcum)
                    dSaldoDeb = Arredonda_Moeda(dSaldoDeb - dSaldoDebAcum)
                    dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
                    dSaldoIniTotal = dSaldoIniTotal + dSaldoIni
                    dSaldoFimTotal = dSaldoFimTotal + dSaldoFim
                
                    If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
                
                        objECD.lSeq = objECD.lSeq + 1
                        lTotalRegBloco1 = lTotalRegBloco1 + 1
                
                        lErro = Mascara_MascararConta(sContaAtual, sContaTela)
                        If lErro <> SUCESSO Then gError 203246
                
                        objECD.sRegistro = "|I155|" & sContaTela & "|" & "|" & Format(Abs(dSaldoIni), "fixed") & "|" & SFF_CD(dSaldoIni) & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & SFF_CD(dSaldoFim) & "|"
                            
                        sContaTelaUltimo = sContaTela
                        dSaldoIniUltimo = dSaldoIni
                        dSaldoFimUltimo = dSaldoFim
                        lSeqUltimo = objECD.lSeq
                        dSaldoCreUltimo = dSaldoCre
                        dSaldoDebUltimo = dSaldoDeb
                        
                        objRegI155.lCount = objRegI155.lCount + 1
                        
                        lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, "", dSaldoIni, objECD, dtDataIniAux, dtDataFimAux, dtDataIni, tPlanoConta.sConta, tPlanoConta.iNaturezaSped)
                        If lErro <> SUCESSO Then gError 207330
                        
                   End If
                     
                End If
            
                lErro = Comando_BuscarProximo(alComando(22))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204747
            
                sContaAtual = tPlanoConta.sConta
                dSaldoAcum = 0
                dSaldoCreAcum = 0
                dSaldoDebAcum = 0
            
            Loop
        
        End If
        
        'deixa 20 locais vagos no final do mes caso tenha que inserir um registro devido a diferenca entre os totais de credito e debito
        objECD.lSeq = objECD.lSeq + 20
        
        If iTipo = SPED_CONTAB_TIPO_FCONT Then Exit Do
        
        dtDataIniAux = DateAdd("m", 1, dtDataIniAux)
        dtDataIniAux = CDate("01/" & Month(dtDataIniAux) & "/" & Year(dtDataIniAux))
        
        dtDataFimAux = DateAdd("m", 1, dtDataIniAux)
        dtDataFimAux = DateAdd("d", -1, dtDataFimAux)
        If dtDataFimAux > dtDataFim Then dtDataFimAux = dtDataFim
    
    Loop
    
    Set colCtaInatSemMov = Nothing
    Set colCclInatSemMov = Nothing
    
    lSeq155Final = objECD.lSeq - 19
    
    If objRegI150.lCount > 0 Then colSpedRegistro.Add objRegI150
    If objRegI155.lCount > 0 Then colSpedRegistro.Add objRegI155
    
    '************ registro I200 *******************
    Call objTela.SPED_Status("Processando Registro I200", "")
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    sCclAux = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(14), "SELECT  FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Data, Conta, Ccl, Historico, Valor FROM Lancamentos WHERE FilialEmpresa < 50 AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) AND Abs(Valor) > 0.001 ORDER BY FilialEmpresa, Origem, Data, Doc, Seq ", _
        tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.sConta, sCclAux, tLanc.sHistorico, tLanc.dValor, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203201
    Else
        lErro = Comando_Executar(alComando(14), "SELECT  FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Data, Conta, Ccl, Historico, Valor FROM Lancamentos WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) AND Abs(Valor) > 0.001 ORDER BY Origem, Data, Doc, Seq ", _
        tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.sConta, sCclAux, tLanc.sHistorico, tLanc.dValor, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203201
    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(14))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203202

    lDocAtual = -99999

    Set objRegI200 = New ClassSpedRegistro
    
    colSpedRegistro.Add objRegI200
    
    objRegI200.sTipo = "I200"

    Set objRegI250 = New ClassSpedRegistro

    objRegI250.sTipo = "I250"

    colSpedRegistro.Add objRegI250

    iFilialEmpresaAtual = tLanc.iFilialEmpresa
    sOrigemAtual = tLanc.sOrigem
    iExercicioAtual = tLanc.iExercicio
    iPeriodoLanAtual = tLanc.iPeriodoLan
    lDocAtual = tLanc.lDoc
    dtDataAtual = tLanc.dtData

    dValorDocPos = 0
    dValorDocNeg = 0
    
    Set colLanc = New Collection

    iX1 = 0
    Do While lErro = AD_SQL_SUCESSO
    
        bSemCcl = False
        For Each vsCta In colCtaSemCcl
            If vsCta = tLanc.sConta Then
                bSemCcl = True
                Exit For
            End If
        Next
        If Not bSemCcl Then
            tLanc.sCcl = sCclAux
        Else
            tLanc.sCcl = ""
        End If
    
        iX1 = iX1 + 1
    
        Call objTela.SPED_Status("Processando Registro I200 - Lançamentos", "Doc: " & CStr(tLanc.lDoc) & " Data: " & Format(tLanc.dtData, "dd/mm/yyyy") & " Valor: " & Format(tLanc.dValor, "STANDARD"))
    
        If iFilialEmpresaAtual <> tLanc.iFilialEmpresa Or sOrigemAtual <> tLanc.sOrigem Or iExercicioAtual <> tLanc.iExercicio Or iPeriodoLanAtual <> tLanc.iPeriodoLan Or lDocAtual <> tLanc.lDoc Then
            
            dValorDoc = dValorDocPos
                
            If dValorDocPos <> dValorDocNeg Then
            
                Set objLanc = colLanc.Item(1)
                
                objLanc.dValor = objLanc.dValor - (dValorDocPos - dValorDocNeg)
                
                If objLanc.dValor > 0 Then
                    dValorDoc = dValorDocNeg
                Else
                    dValorDoc = dValorDocPos
                End If
                
            End If
            
            sIdentificador = Format(iFilialEmpresaAtual, "00") & sOrigemAtual & Format(iExercicioAtual, "00") & Format(iPeriodoLanAtual, "00") & Format(lDocAtual, "000000000")
            
            objECD.sTipo = "I200"
            objECD.lSeq = objECD.lSeq + 1
            lTotalRegBloco1 = lTotalRegBloco1 + 1
            objECD.sRegistro = "|I200|" & sIdentificador & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & IIf(sOrigemAtual = "APE" Or sOrigemAtual = "EAE", "E", "N") & "|"
            If iCodVersao >= 7 Then objECD.sRegistro = objECD.sRegistro & "|"   '6 DT_LCTO_EXT
            objRegI200.lCount = objRegI200.lCount + 1
            
            If Len(objECD.sRegistro) > 500 Then gError 203205
            
            If iTipo = SPED_CONTAB_TIPO_NORMAL Then
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                If lErro <> SUCESSO Then gError 203206
            End If
            
            For Each objLanc In colLanc
                
                objECD.sTipo = "I250"
                objECD.lSeq = objECD.lSeq + 1
                lTotalRegBloco1 = lTotalRegBloco1 + 1
            
                objRegI250.lCount = objRegI250.lCount + 1
                
                objECD.sRegistro = "|I250|" & objLanc.sConta & "|" & objLanc.sCcl & "|" & Format(Abs(objLanc.dValor), "fixed") & "|" & SFF_CD(objLanc.dValor) & "|||" & SFF_Texto(objLanc.sHistorico) & "||"
            
                If Len(objECD.sRegistro) > 500 Then gError 203209
        
                If iTipo = SPED_CONTAB_TIPO_NORMAL Then
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                    If lErro <> SUCESSO Then gError 203210
                End If
                
                lErro = SpedContabil1_Atualiza_CredDeb(colSpedContabil1, objLanc.sConta, objLanc.sCcl, objLanc.dValor, objLanc.dtData, iTipo)
                If lErro <> SUCESSO Then gError 207331
            
            Next
            
            iFilialEmpresaAtual = tLanc.iFilialEmpresa
            sOrigemAtual = tLanc.sOrigem
            iExercicioAtual = tLanc.iExercicio
            iPeriodoLanAtual = tLanc.iPeriodoLan
            lDocAtual = tLanc.lDoc
            dtDataAtual = tLanc.dtData
            
            dValorDocPos = 0
            dValorDocNeg = 0
            
            Set colLanc = New Collection
            
        End If
    
        '************ registro I250 *******************
         Call objTela.SPED_Status("Processando Registro I250 - Lançamentos", "Doc: " & CStr(tLanc.lDoc) & " Data: " & Format(tLanc.dtData, "dd/mm/yyyy") & " Valor: " & Format(tLanc.dValor, "STANDARD"))
        
        lErro = Mascara_MascararConta(tLanc.sConta, sContaTela)
        If lErro <> SUCESSO Then gError 203207
        
        'se usa cccusto extra contabil
        If iUsoCcl = 2 And Len(Trim(tLanc.sCcl)) > 0 Then
            lErro = Mascara_MascararCcl(tLanc.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 203208
        Else
            sCclMascarado = ""
        End If
        
        If tLanc.dValor > 0 Then dValorDocPos = dValorDocPos + CDbl(Format(Abs(Arredonda_Moeda(tLanc.dValor)), "fixed"))
        If tLanc.dValor < 0 Then dValorDocNeg = dValorDocNeg + CDbl(Format(Abs(Arredonda_Moeda(tLanc.dValor)), "fixed"))
       
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.sConta = sContaTela
        objLanc.sCcl = sCclMascarado
        objLanc.dValor = Arredonda_Moeda(tLanc.dValor)
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dtData = tLanc.dtData
        
        colLanc.Add objLanc
        
        'verifica se o lancamento envolve uma conta de resultado, se envolver guarda o valor
        For Each objSpedContRes1 In colSpedContRes1
            'se for um lancamento que envolve conta de resultado
            If tLanc.sConta = objSpedContRes1.sConta And (tLanc.sOrigem = "APE" Or tLanc.sOrigem = "EAE") Then
            
                iAchou2 = 0
                
                For Each objSpedContRes2 In colSpedContRes2
                    'se a conta de resultado ja estiver armazenada, acumula seu valor
                    If objSpedContRes2.sConta = tLanc.sConta And objSpedContRes2.sCcl = tLanc.sCcl Then
                        iAchou2 = 1
                        objSpedContRes2.dCred = objSpedContRes2.dCred + Arredonda_Moeda(tLanc.dValor)
                        Exit For
                    End If
                Next
                    
                'se nao estava cadastrada a conta/ccl ainda, cadastra neste momento
                If iAchou2 = 0 Then
                
                    Set objSpedContRes2 = New ClassSpedContabil1
                    
                    objSpedContRes2.dtDataIni = tLanc.dtData
                    objSpedContRes2.sConta = tLanc.sConta
                    objSpedContRes2.sCcl = tLanc.sCcl
                    objSpedContRes2.dCred = Arredonda_Moeda(tLanc.dValor)
                    
                    colSpedContRes2.Add objSpedContRes2
                    
                End If
        
            End If
        
        Next
    
        objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
        objTela.dValorAtual = objTela.dValorAtual + 1
    
        If objTela.dValorAtual < objTela.dValorTotal Then objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
    
        DoEvents
        
        If objTela.iCancelaBatch = CANCELA_BATCH Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
        
            If vbMsgRes = vbYes Then gError 203264
            
            objTela.iCancelaBatch = 0
                
        End If
    
        lErro = Comando_BuscarProximo(alComando(14))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203211
    
    Loop
    
    dValorDoc = dValorDocPos
    
    If dValorDocPos <> dValorDocNeg Then
    
        Set objLanc = colLanc.Item(1)
        
        objLanc.dValor = objLanc.dValor - (dValorDocPos - dValorDocNeg)
        
        If objLanc.dValor > 0 Then
            dValorDoc = dValorDocNeg
        Else
            dValorDoc = dValorDocPos
        End If
        
        
    End If
        
    If colLanc.Count > 0 Then
        
        sIdentificador = Format(iFilialEmpresaAtual, "00") & sOrigemAtual & Format(iExercicioAtual, "00") & Format(iPeriodoLanAtual, "00") & Format(lDocAtual, "000000000")
        
        objECD.sTipo = "I200"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objECD.sRegistro = "|I200|" & sIdentificador & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & IIf(sOrigemAtual = "APE" Or sOrigemAtual = "EAE", "E", "N") & "|"
        If iCodVersao >= 7 Then objECD.sRegistro = objECD.sRegistro & "|"   '6 DT_LCTO_EXT

        objRegI200.lCount = objRegI200.lCount + 1
        
        If Len(objECD.sRegistro) > 500 Then gError 203205
        
        If iTipo = SPED_CONTAB_TIPO_NORMAL Then
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203206
        End If
    
    End If
    
    For Each objLanc In colLanc
        
        objECD.sTipo = "I250"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objRegI250.lCount = objRegI250.lCount + 1
    
        objECD.sRegistro = "|I250|" & objLanc.sConta & "|" & objLanc.sCcl & "|" & Format(Abs(objLanc.dValor), "fixed") & "|" & SFF_CD(objLanc.dValor) & "|||" & objLanc.sHistorico & "||"
    
        If Len(objECD.sRegistro) > 500 Then gError 203209

        If iTipo = SPED_CONTAB_TIPO_NORMAL Then
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203210
        End If
        
        lErro = SpedContabil1_Atualiza_CredDeb(colSpedContabil1, objLanc.sConta, objLanc.sCcl, objLanc.dValor, objLanc.dtData, iTipo)
        If lErro <> SUCESSO Then gError 207331
    
    Next
    
    iContador = 0
    
    Set colSpedContabil1_Lanc = New Collection
        
    For Each objECD1 In colSpedContabil1
    
        iContador = iContador + 1
    
        Set objECDAux1 = objECD1.objSpedContabil
        
        dSaldoFim = objECD1.dSldIni + objECD1.dCred - objECD1.dDeb
        
        'se se tratar do mes de encerramento e tiver saldoFim ==> todas as contas de encerramentos tem q ter o saldo final zerados
        'se nao for o acaso, vamos criar um lancamento com a primeira conta de ativo e as demais contas de encerramento
        'com a diferenca
        If objECD1.dtDataIni = CDate("01/12/" & Year(objECD1.dtDataIni)) And Abs(dSaldoFim) >= 0.01 Then
            
            tPlanoConta.sConta = String(STRING_CONTA, 0)
            tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
            
            lErro = Comando_Executar(alComando(3), "SELECT Conta, TipoConta, NivelConta, DescConta, NaturezaSped FROM PlanoConta WHERE Conta = ?", _
            tPlanoConta.sConta, tPlanoConta.iTipoConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, iNaturezaSped, objECD1.sContaBD)
            If lErro <> AD_SQL_SUCESSO Then gError 203139
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203140
        
            'se for conta de resultado ==> guarda os dados para criar um lancamento zerando o saldo
            If iNaturezaSped = 4 Then
                
                'guarda os dados para criar o lancamento
                Set objECD1_Lanc = New ClassSpedContabil1
                
                objECD1_Lanc.dCred = objECD1.dCred
                objECD1_Lanc.dDeb = objECD1.dDeb
                objECD1_Lanc.dSldIni = objECD1.dSldIni
                objECD1_Lanc.sContaBD = objECD1.sContaBD
                objECD1_Lanc.sConta = objECD1.sConta
                objECD1_Lanc.sCcl = objECD1.sCcl
                objECD1_Lanc.dtDataIni = objECD1.dtDataIni
                                
                colSpedContabil1_Lanc.Add objECD1_Lanc
                
                'acerta o credito ou debito para zerar o saldo final
                If dSaldoFim > 0 Then
                    objECD1.dDeb = objECD1.dDeb + Arredonda_Moeda(dSaldoFim)
                Else
                    objECD1.dCred = objECD1.dCred + Arredonda_Moeda(Abs(dSaldoFim))
                End If
                
                dSaldoFim = 0
                
            End If
        
        End If
        
        If objECD1.dSldIni > 0 Then
            dTotalSldIniCred = dTotalSldIniCred + objECD1.dSldIni
        Else
            dTotalSldIniDeb = dTotalSldIniDeb + Abs(objECD1.dSldIni)
        End If
        
        iAchou = 0
        
        If iContador + 1 <= colSpedContabil1.Count Then
            Set objECD1_Prox = colSpedContabil1(iContador + 1)
            If objECD1_Prox.dtDataIni > objECD1.dtDataIni Then
                iAchou = 1
            End If
        End If
        
'        If (iContador = colSpedContabil1.Count Or iAchou = 1) And objECD1.dtDataIni = dtDataIni Then
        If (iContador = colSpedContabil1.Count Or iAchou = 1) Then
            objECD1.dSldIni = objECD1.dSldIni - (dTotalSldIniCred - dTotalSldIniDeb)
            dSaldoFim = dSaldoFim - (dTotalSldIniCred - dTotalSldIniDeb)
            dDiferenca = dTotalSldIniCred - dTotalSldIniDeb
            dTotalSldIniCred = 0
            dTotalSldIniDeb = 0
        End If
        
        'colocar o saldo final como o saldo inicial do proximo periodo
        lErro = SpedContabil1_SaldoFinal_SaldoInicial(colSpedContabil1, objECD1.sConta, objECD1.sCcl, dSaldoFim, objECD1.dtDataIni, iAchou1)
        If lErro <> SUCESSO Then gError 207353
        
        If Arredonda_Moeda(objECD1.dSldIni) <> 0 Or Arredonda_Moeda(objECD1.dDeb) <> 0 Or Arredonda_Moeda(objECD1.dCred) <> 0 Or Arredonda_Moeda(dSaldoFim) <> 0 Then
        
            objECDAux1.sRegistro = "|I155|" & objECD1.sConta & "|" & objECD1.sCcl & "|" & Format(Abs(Arredonda_Moeda(objECD1.dSldIni)), "fixed") & "|" & SFF_CD(objECD1.dSldIni) & "|" & Format(objECD1.dDeb, "fixed") & "|" & Format(objECD1.dCred, "fixed") & "|" & Format(Abs(Arredonda_Moeda(dSaldoFim)), "fixed") & "|" & SFF_CD(dSaldoFim) & "|"
                    
            If Len(objECD.sRegistro) > 500 Then gError 203240
            
            For Each vRegM025Aux In colM025Aux
                If left(vRegM025Aux, Len("|M025|" & objECD1.sConta & "|" & objECD1.sCcl & "|")) = "|M025|" & objECD1.sConta & "|" & objECD1.sCcl & "|" Then
                    vRegM025 = vRegM025Aux & Format(Abs(objECD1.dSldIni), "fixed") & "|" & SFF_CD(objECD1.dSldIni) & "|" & Format(Abs(objECD1.dSldIni), "fixed") & "|" & SFF_CD(objECD1.dSldIni) & "|"
                    colM025.Add vRegM025
                    Exit For
                End If
            Next
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECDAux1.lNumIntSped, objECDAux1.lSeq, objECDAux1.iFilialEmpresa, objECDAux1.sTipo, objECDAux1.sRegistro)
            If lErro <> SUCESSO Then gError 203241
            
            If objECD1.dtDataIni = CDate("01/12/" & Year(objECD1.dtDataIni)) And objECD1.sContaBD = sContaOutros Then
                Set objECD1_Outros = objECD1
            End If
            
        Else
            lTotalRegBloco1 = lTotalRegBloco1 - 1
            objRegI155.lCount = objRegI155.lCount - 1
        End If
    
        If iAchou1 = 0 And Abs(dDiferenca) >= 0.01 Then
        
            lErro = SpedContabil1_SaldoInicial(colSpedContabil1, objECD1.sConta, objECD1.sCcl, dDiferenca, objECD1.dtDataIni, dtDataFim, objECD1.sContaBD, lTotalRegBloco1, objRegI155, objECD1.iNaturezaSped)
            If lErro <> SUCESSO Then gError 207353
            
            dDiferenca = 0
        End If
    Next
    
    dValorDocPos = 0
    dValorDocNeg = 0
    
    Set colLanc = New Collection
    
    'cria um lancamento capaz de zerar as contas de encerramento causa tenha sobrado algum residuo devido ao arredondamento dos lancamentos
    For Each objECD1_Lanc In colSpedContabil1_Lanc
        
        dSaldoFim = objECD1_Lanc.dSldIni + objECD1_Lanc.dCred - objECD1_Lanc.dDeb
        
        If -dSaldoFim > 0 Then
            dValorDocPos = dValorDocPos - dSaldoFim
        Else
            dValorDocNeg = dValorDocNeg + dSaldoFim
        End If
        
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.dtData = CDate("31/12/" & Year(dtDataIni))
        objLanc.sConta = objECD1_Lanc.sConta
        objLanc.sCcl = objECD1_Lanc.sCcl
        objLanc.dValor = -dSaldoFim
        objLanc.sHistorico = "*****"
        
        colLanc.Add objLanc
    
    Next
    
    'cria um lancamento capaz de zerar as contas de encerramento causa tenha sobrado algum residuo devido ao arredondamento dos lancamentos
    If colLanc.Count > 0 Then
    
        If dValorDocPos > dValorDocNeg Then
            dValorDoc = dValorDocPos
        Else
            dValorDoc = dValorDocNeg
        End If
    
        lDocAtual = 999999999
        dtDataAtual = CDate("31/12/" & Year(dtDataIni))
    
        sIdentificador = Format(iFilialEmpresaAtual, "00") & "SPD" & Format(iExercicioAtual, "00") & "12" & Format(lDocAtual, "000000000")
            
        objECD.sTipo = "I200"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objECD.sRegistro = "|I200|" & sIdentificador & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & "N" & "|"
        
        objRegI200.lCount = objRegI200.lCount + 1
        
        If Len(objECD.sRegistro) > 500 Then gError 203205
        
        If iTipo = SPED_CONTAB_TIPO_NORMAL Then
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203206
        End If
        
        For Each objLanc In colLanc
            
            objECD.sTipo = "I250"
            objECD.lSeq = objECD.lSeq + 1
            lTotalRegBloco1 = lTotalRegBloco1 + 1
            objRegI250.lCount = objRegI250.lCount + 1
        
            objECD.sRegistro = "|I250|" & objLanc.sConta & "|" & objLanc.sCcl & "|" & Format(Abs(objLanc.dValor), "fixed") & "|" & SFF_CD(objLanc.dValor) & "|||" & objLanc.sHistorico & "||"
        
            If Len(objECD.sRegistro) > 500 Then gError 203209
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203210
        
        Next
    
        objECD.sTipo = "I250"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objRegI250.lCount = objRegI250.lCount + 1
    
        dValor = -(dValorDocPos - dValorDocNeg)
        
        If Abs(dValor) > DELTA_VALORMONETARIO Then
        
            If Len(Trim(sContaOutros)) = 0 Then gError 211569
        
            lErro = Mascara_MascararConta(sContaOutros, sContaTela)
            If lErro <> SUCESSO Then gError 203207
            
            objECD.sRegistro = "|I250|" & sContaTela & "|" & "|" & Format(Abs(dValor), "fixed") & "|" & SFF_CD(dValor) & "|||" & "*****" & "||"
        
            If Len(objECD.sRegistro) > 500 Then gError 203209
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203210
        
            If objECD1_Outros Is Nothing Then
            
                Set objECDAux1 = New ClassSpedContabil
        
                objECDAux1.iFilialEmpresa = objECD.iFilialEmpresa
                objECDAux1.lNumIntSped = objECD.lNumIntSped
                objECDAux1.sTipo = "I155"
                
                lErro = Mascara_MascararConta(sContaOutros, sContaTela)
                If lErro <> SUCESSO Then gError 203207
            
                objECDAux1.sRegistro = "|I155|" & sContaTela & "|" & "|" & "0,00" & "|" & "C" & "|" & Format(IIf(dValor < 0, Abs(dValor), 0), "fixed") & "|" & Format(IIf(dValor > 0, dValor, 0), "fixed") & "|" & Format(Abs(dValor), "fixed") & "|" & SFF_CD(dValor) & "|"
        
                For Each vRegM025Aux In colM025Aux
                    If left(vRegM025Aux, Len("|M025|" & sContaTela & "||")) = "|M025|" & sContaTela & "||" Then
                        vRegM025 = vRegM025Aux & "0,00|C|0,00|C|"
                        colM025.Add vRegM025
                        Exit For
                    End If
                Next
        
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objECDAux1.lNumIntSped, lSeq155Final, objECDAux1.iFilialEmpresa, objECDAux1.sTipo, objECDAux1.sRegistro)
                If lErro <> SUCESSO Then gError 203241
                
                lTotalRegBloco1 = lTotalRegBloco1 + 1
                objRegI155.lCount = objRegI155.lCount + 1
        
            Else
            
                If dValor < 0 Then
                    objECD1_Outros.dDeb = objECD1_Outros.dDeb + Abs(dValor)
                Else
                    objECD1_Outros.dCred = objECD1_Outros.dCred + dValor
                End If
        
                dSaldoFim = objECD1_Outros.dSldIni + objECD1_Outros.dCred - objECD1_Outros.dDeb
        
                objECD.sRegistro = "|I155|" & objECD1_Outros.sConta & "|" & "|" & Format(Abs(Arredonda_Moeda(objECD1_Outros.dSldIni)), "fixed") & "|" & SFF_CD(objECD1_Outros.dSldIni) & "|" & Format(objECD1_Outros.dDeb, "fixed") & "|" & Format(objECD1_Outros.dCred, "fixed") & "|" & Format(Abs(Arredonda_Moeda(dSaldoFim)), "fixed") & "|" & SFF_CD(dSaldoFim) & "|"
        
                For iIndice = 1 To colM025.Count
                    vRegM025 = colM025.Item(iIndice)
                    If left(vRegM025, Len("|M025|" & objECD1_Outros.sConta & "|" & objECD1_Outros.sCcl & "|")) = "|M025|" & objECD1_Outros.sConta & "|" & objECD1_Outros.sCcl & "|" Then
                        colM025.Remove iIndice
                        Exit For
                    End If
                Next
        
                For Each vRegM025Aux In colM025Aux
                    If left(vRegM025Aux, Len("|M025|" & objECD1_Outros.sConta & "|" & objECD1_Outros.sCcl & "|")) = "|M025|" & objECD1_Outros.sConta & "|" & objECD1_Outros.sCcl & "|" Then
                        vRegM025 = vRegM025Aux & Format(Abs(objECD1_Outros.dSldIni), "fixed") & "|" & SFF_CD(objECD1_Outros.dSldIni) & "|" & Format(Abs(objECD1_Outros.dSldIni), "fixed") & "|" & SFF_CD(objECD1_Outros.dSldIni) & "|"
                        colM025.Add vRegM025
                        Exit For
                    End If
                Next
        
                lErro = Comando_Executar(alComando(1), "UPDATE SpedContabil SET  Registro = ? WHERE NumIntSped = ? AND Seq = ? AND FilialEmpresa = ? AND Tipo = ?", _
                objECD.sRegistro, objECD1_Outros.objSpedContabil.lNumIntSped, objECD1_Outros.objSpedContabil.lSeq, objECD1_Outros.objSpedContabil.iFilialEmpresa, objECD1_Outros.objSpedContabil.sTipo)
                If lErro <> SUCESSO Then gError 203210
            
            End If
                
        End If
        
    End If
    
    Set colM025Aux = Nothing
    
    '************ registro I350 *******************
    If colSpedContRes2.Count > 0 Then
        Set objSpedContRes2 = colSpedContRes2.Item(1)
                
        objECD.sTipo = "I350"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objECD.sRegistro = "|I350|" & Format(objSpedContRes2.dtDataIni, "ddmmyyyy") & "|"
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        If Len(objECD.sRegistro) > 500 Then gError 207354
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 207355
    End If
    
    '************ registro I355 *******************
    If colSpedContRes2.Count > 0 Then
        Set objReg = New ClassSpedRegistro
        colSpedRegistro.Add objReg
    End If
    
    dLucro = 0
    For Each objSpedContRes2 In colSpedContRes2
    
        lErro = Mascara_MascararConta(objSpedContRes2.sConta, sContaTela)
        If lErro <> SUCESSO Then gError 207356
        
        'se usa cccusto extra contabil
        If iUsoCcl = 2 And Len(Trim(objSpedContRes2.sCcl)) > 0 Then
            lErro = Mascara_MascararCcl(objSpedContRes2.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 207357
        Else
            sCclMascarado = ""
        End If
    
        objECD.sTipo = "I355"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
        objECD.sRegistro = "|I355|" & sContaTela & "|" & sCclMascarado & "|" & Format(Abs(Arredonda_Moeda(objSpedContRes2.dCred)), "fixed") & "|" & SFF_CD(objSpedContRes2.dCred * -1) & "|"
        
        dLucro = dLucro + Arredonda_Moeda(objSpedContRes2.dCred)
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        If Len(objECD.sRegistro) > 500 Then gError 207358
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 207359
    
    Next
    
    '************ registro I990 *******************
    Call objTela.SPED_Status("Processando Registro I990", "")

    objECD.sTipo = "I990"
    objECD.lSeq = objECD.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    objECD.sRegistro = "|I990|" & lTotalRegBloco1 & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203212
    
    '************ registro J001 *******************
    Call objTela.SPED_Status("Processando Registro J001", "")
  
    objECD.sTipo = "J001"
    objECD.lSeq = objECD.lSeq + 1
    lTotalRegBlocoJ = lTotalRegBlocoJ + 1
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    objECD.sRegistro = "|J001|0|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203213
    
    If iTipo = SPED_CONTAB_TIPO_NORMAL Then
        
        '************ registro J005 *******************
        Call objTela.SPED_Status("Processando Registro J005", "")
      
        objECD.sTipo = "J005"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoJ = lTotalRegBlocoJ + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|J005|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|1||"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 207332
        
        '************ registro J100 *******************
        tPlanoConta.sConta = String(STRING_CONTA, 0)
        tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
        objECD.sTipo = "J100"
    
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
    
        lErro = Comando_Executar(alComando(31), "SELECT Conta, NivelConta, DescConta, NaturezaSped, Categoria, TipoConta FROM PlanoConta WHERE NaturezaSped IN (1,2,3) ORDER BY Conta", _
        tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, iNaturezaSped, tPlanoConta.iCategoria, tPlanoConta.iTipoConta)
        If lErro <> AD_SQL_SUCESSO Then gError 207333
    
        lErro = Comando_BuscarPrimeiro(alComando(31))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207334
        
        Do While lErro = AD_SQL_SUCESSO
        
            Call Mascara_RetornaContaEnxuta(tPlanoConta.sConta, sContaEnxuta)
    
            dValor = 0
            For Each objECD1 In colSpedContabil1
                If objECD1.dtDataFim = dtDataFim And left(objECD1.sContaBD, Len(sContaEnxuta)) = sContaEnxuta And (objECD1.iNaturezaSped = 1 Or objECD1.iNaturezaSped = 2 Or objECD1.iNaturezaSped = 3) Then
                    dValor = Arredonda_Moeda(dValor + objECD1.dSldIni + objECD1.dCred - objECD1.dDeb)
                End If
            Next
                       
            dValorIni = 0
            For Each objECD1 In colSpedContabil1
                If objECD1.dtDataIni = dtDataIni And left(objECD1.sContaBD, Len(sContaEnxuta)) = sContaEnxuta And (objECD1.iNaturezaSped = 1 Or objECD1.iNaturezaSped = 2 Or objECD1.iNaturezaSped = 3) Then
                    dValorIni = Arredonda_Moeda(dValorIni + objECD1.dSldIni)
                End If
            Next
            
            If Abs(dValor) > DELTA_VALORMONETARIO Or Abs(dValorIni) > DELTA_VALORMONETARIO Then
            
                If tPlanoConta.iNivelConta > 1 Then
                    
                    lErro = Mascara_RetornaContaPai(tPlanoConta.sConta, sConta)
                    If lErro <> SUCESSO Then gError 203144
                
                    lErro = Mascara_MascararConta(sConta, sContaPai)
                    If lErro <> SUCESSO Then gError 203143
                
                Else
                    sContaPai = ""
                End If
                 
                objECD.lSeq = objECD.lSeq + 1
                
                objReg.lCount = objReg.lCount + 1
                lTotalRegBlocoJ = lTotalRegBlocoJ + 1
                
                lErro = Mascara_MascararConta(tPlanoConta.sConta, sContaTela)
                If lErro <> SUCESSO Then gError 203207
                
                objECD.sRegistro = "|J100|" '01-REG
                objECD.sRegistro = objECD.sRegistro & sContaTela & "|" '02-COD_AGL
                
                If tPlanoConta.iTipoConta = 1 Then
                    objECD.sRegistro = objECD.sRegistro & IIf(iCodVersao >= 7, "T|", "") '03-IND_COD_AGL
                Else
                    objECD.sRegistro = objECD.sRegistro & IIf(iCodVersao >= 7, "D|", "") '03-IND_COD_AGL
                End If
                objECD.sRegistro = objECD.sRegistro & tPlanoConta.iNivelConta & "|" '04-NIVEL_AGL
                objECD.sRegistro = objECD.sRegistro & IIf(iCodVersao >= 7, sContaPai & "|", "") '05-COD_AGL_SUP
                If iCodVersao >= 7 Then
                    objECD.sRegistro = objECD.sRegistro & IIf(iNaturezaSped = 1, "A", "P") & "|" '06-IND_GRP_BAL
                Else
                    objECD.sRegistro = objECD.sRegistro & IIf(iNaturezaSped = 1, "1", "2") & "|" '06-IND_GRP_BAL
                End If
                objECD.sRegistro = objECD.sRegistro & tPlanoConta.sDescConta & "|" '07-DESCR_COD_AGL
                If iCodVersao >= 7 Then
                    objECD.sRegistro = objECD.sRegistro & Abs(dValorIni) & "|" '08-VL_CTA_INI
                    objECD.sRegistro = objECD.sRegistro & SFF_CD(dValorIni) & "|" '09-IND_DC_CTA_INI
                    objECD.sRegistro = objECD.sRegistro & Abs(dValor) & "|" '10-VL_CTA_FIN
                    objECD.sRegistro = objECD.sRegistro & SFF_CD(dValor) & "|" '11-IND_DC_CTA_FIN
               Else
                    objECD.sRegistro = objECD.sRegistro & Abs(dValor) & "|"
                    objECD.sRegistro = objECD.sRegistro & SFF_CD(dValor) & "|"
                    If iCodVersao >= 2 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & Abs(dValorIni) & "|" & SFF_CD(dValorIni) & "|"
                End If
                If iCodVersao >= 6 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & "|"  '12-NOTA_EXP_REF
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
                If lErro <> SUCESSO Then gError 207336
                
            End If
        
            lErro = Comando_BuscarProximo(alComando(31))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207337
        
        Loop
        
        '************ registro J150 *******************
        lErro = CF("RelDRE_Calcula", "SPED", objPer.iExercicio, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 207338
    
        objECD.sTipo = "J150"
    
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
    
        tRelDRE.sTitulo = String(255, 0)
    
        lErro = Comando_Executar(alComando(32), "SELECT Codigo, Tipo, Nivel, Titulo, Valor, ValorExercAnt FROM RelDRE WHERE Relatorio = 'DRE' AND Modelo = 'SPED' AND Tipo IN (0,1) ORDER BY Posicao", _
        tRelDRE.iCodigo, tRelDRE.iTipo, tRelDRE.iNivel, tRelDRE.sTitulo, tRelDRE.dValor, dValorExercAnt)
        If lErro <> AD_SQL_SUCESSO Then gError 207339
    
        lErro = Comando_BuscarPrimeiro(alComando(32))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207340
        
        Do While lErro = AD_SQL_SUCESSO
        
            objECD.lSeq = objECD.lSeq + 1
            
            objReg.lCount = objReg.lCount + 1
            lTotalRegBlocoJ = lTotalRegBlocoJ + 1
            
            objECD.sRegistro = "|J150|"            '01 REG
            objECD.sRegistro = objECD.sRegistro & tRelDRE.iCodigo & "|" '02 COD_AGL
            objECD.sRegistro = objECD.sRegistro & IIf(iCodVersao >= 7, "T|", "") '03 IND_COD_AGL
            objECD.sRegistro = objECD.sRegistro & tRelDRE.iNivel & "|" '04 NIVEL_AGL
            objECD.sRegistro = objECD.sRegistro & IIf(iCodVersao >= 7, "|", "") '05 COD_AGL_SUP
            objECD.sRegistro = objECD.sRegistro & tRelDRE.sTitulo & "|" '06 DESCR_COD_AGL
            objECD.sRegistro = objECD.sRegistro & Arredonda_Moeda(Abs(tRelDRE.dValor)) & "|" '07 VL_CTA

            If iCodVersao >= 7 Then
                objECD.sRegistro = objECD.sRegistro & IIf(tRelDRE.dValor > 0, "C", "D") & "|"  '08 IND_DC_CTA
                'objECD.sRegistro = objECD.sRegistro & Arredonda_Moeda(Abs(dValorExercAnt)) & "|"
                objECD.sRegistro = objECD.sRegistro & IIf(dValorExercAnt > 0, "R", "D") & "|"  '09 IND_GRP_DRE
            Else
                objECD.sRegistro = objECD.sRegistro & IIf(tRelDRE.iTipo = 0, IIf(tRelDRE.dValor > 0, "R", "D"), IIf(tRelDRE.dValor > 0, "P", "N")) & "|"
                
                If iCodVersao >= 4 Then
                    objECD.sRegistro = objECD.sRegistro & Arredonda_Moeda(Abs(dValorExercAnt)) & "|"
                    objECD.sRegistro = objECD.sRegistro & IIf(tRelDRE.iTipo = 0, IIf(dValorExercAnt > 0, "R", "D"), IIf(dValorExercAnt > 0, "P", "N")) & "|"
                End If
            End If
            
            If iCodVersao >= 6 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & "|"  '10 NOTA_EXP_REF
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 207341
        
            lErro = Comando_BuscarPrimeiro(alComando(32))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207342
        
        Loop
        
        '************ registro J900 *******************
        Call objTela.SPED_Status("Processando Registro J900", "")
    
        objECD.sTipo = "J900"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoJ = lTotalRegBlocoJ + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|J900|TERMO DE ENCERRAMENTO|" & lNumOrd & "|Escrituração Contábil Digital do Livro Diário Geral|"
        objECD.sRegistro = objECD.sRegistro & gsNomeEmpresa & "||" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
        
        If Len(objECD.sRegistro) > 500 Then gError 203214
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203215
        
    End If
        
    '************ registro J930 *******************
    Call objTela.SPED_Status("Processando Registro J930", "")

    objECD.sTipo = "J930"
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    lTotalRegBlocoJ = lTotalRegBlocoJ + 1
    
    sDescQualiSigCTB = String(STRING_DNRC_DESC, 0)
    
    lErro = Comando_Executar(alComando(15), "SELECT  Descricao FROM DNRC WHERE Codigo = ?", _
    sDescQualiSigCTB, objFilialEmpresa.sCodQualiSigCTB)
    If lErro <> AD_SQL_SUCESSO Then gError 203216
    
    lErro = Comando_BuscarPrimeiro(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203217
    
    objECD.sRegistro = "|J930|" & objFilialEmpresa.sSignatarioCTB & "|" & objFilialEmpresa.sCPFSignatarioCTB & "|" & sDescQualiSigCTB & "|" & objFilialEmpresa.sCodQualiSigCTB & "||"
    
    If iCodVersao >= 2 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & "|||||"
    If iCodVersao >= 3 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & "S|"
    
    objECD.lSeq = objECD.lSeq + 1
    
    If Len(objECD.sRegistro) > 500 Then gError 203218
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203219
    
    objReg.lCount = objReg.lCount + 1
    objECD.lSeq = objECD.lSeq + 1
    lTotalRegBlocoJ = lTotalRegBlocoJ + 1
    
    objECD.sRegistro = "|J930|" & objFilialEmpresa.sContador & "|" & objFilialEmpresa.sCPFContador & "|" & "Contador" & "|" & "900" & "|" & objFilialEmpresa.sCRCContador & "|"
    
    If iCodVersao >= 2 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & objFilialEmpresa.objEnderecoContador.sEmail & "|" & Replace(Replace(objFilialEmpresa.objEnderecoContador.sTelefone1, " ", ""), "-", "") & "|" & objFilialEmpresa.objEnderecoContador.sSiglaEstado & "|||"
    If iCodVersao >= 3 And iTipo = SPED_CONTAB_TIPO_NORMAL Then objECD.sRegistro = objECD.sRegistro & "N|"
    
    If Len(objECD.sRegistro) > 500 Then gError 203248
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203249
    
    
    '************ registro J935 *******************
    If iCodVersao >= 3 And iTipo = SPED_CONTAB_TIPO_NORMAL And iEmpGrandePorte = 1 Then
    
        Call objTela.SPED_Status("Processando Registro J935", "")
    
        objECD.sTipo = "J935"
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        lTotalRegBlocoJ = lTotalRegBlocoJ + 1
           
        objECD.sRegistro = "|J935|||"
        If iCodVersao >= 7 Then objECD.sRegistro = objECD.sRegistro + "|"
        
        objECD.lSeq = objECD.lSeq + 1
        
        If Len(objECD.sRegistro) > 500 Then gError 203218
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203219
    
    End If
    
    '************ registro J990 *******************
    Call objTela.SPED_Status("Processando Registro J990", "")
    
    objECD.sTipo = "J990"
    objECD.lSeq = objECD.lSeq + 1
    lTotalRegBlocoJ = lTotalRegBlocoJ + 1
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    objECD.sRegistro = "|J990|" & lTotalRegBlocoJ & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203220
    
    If iTipo = SPED_CONTAB_TIPO_FCONT Then
        'REGISTROS M
        
        '************ registro M001 *******************
        Call objTela.SPED_Status("Processando Registro M001", "")
        
        objECD.sTipo = "M001"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoM = lTotalRegBlocoM + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|M001|0|"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203220
        
        '************ registro M020 *******************
        Call objTela.SPED_Status("Processando Registro M020", "")
        
        objECD.sTipo = "M020"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoM = lTotalRegBlocoM + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        'Tem que ver de onde pegar todas essas configurações
        objECD.sRegistro = "|M020|10|0|||||A|1|||"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203220
        
        '************ registro M025 *******************
        
        Call objTela.SPED_Status("Processando Registro M025", "")
        objECD.sTipo = "M025"
        Set objReg = New ClassSpedRegistro
        colSpedRegistro.Add objReg
        objReg.sTipo = objECD.sTipo
        
        For Each vRegM025 In colM025
        
            objECD.lSeq = objECD.lSeq + 1
            lTotalRegBlocoM = lTotalRegBlocoM + 1
            
            objReg.lCount = objReg.lCount + 1
            
            'Tem que ver de onde pegar todas essas configurações
            objECD.sRegistro = vRegM025
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
            If lErro <> SUCESSO Then gError 203220
        
        Next
        
        Set colM025 = Nothing
        
        '************ registro M030 *******************
        Call objTela.SPED_Status("Processando Registro M030", "")
        
        objECD.sTipo = "M030"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoM = lTotalRegBlocoM + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|M030|A00|" & Format(Abs(Arredonda_Moeda(dLucro)), "fixed") & "|" & SFF_CD(dLucro * -1) & "|"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203220
        
        '************ registro M990 *******************
        Call objTela.SPED_Status("Processando Registro M990", "")
        
        objECD.sTipo = "M990"
        objECD.lSeq = objECD.lSeq + 1
        lTotalRegBlocoM = lTotalRegBlocoM + 1
        
        Set objReg = New ClassSpedRegistro
        
        colSpedRegistro.Add objReg
        
        objReg.sTipo = objECD.sTipo
        objReg.lCount = objReg.lCount + 1
        
        objECD.sRegistro = "|M990|" & lTotalRegBlocoM & "|"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203220


    End If

    '************ registro 9001 *******************
    Call objTela.SPED_Status("Processando Registro 9001", "")
   
    objECD.sTipo = "9001"
    objECD.lSeq = objECD.lSeq + 1
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = objECD.sTipo
    objReg.lCount = objReg.lCount + 1
    
    objECD.sRegistro = "|9001|0|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203221

    '************ registro 9900 *******************
    Call objTela.SPED_Status("Processando Registro 9900", "")
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = "9900"
    objReg.lCount = colSpedRegistro.Count + 2
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = "9990"
    objReg.lCount = 1
    
    Set objReg = New ClassSpedRegistro
    
    colSpedRegistro.Add objReg
    
    objReg.sTipo = "9999"
    objReg.lCount = 1
    
    objECD.sTipo = "9900"
    
    iX1 = 0
    For Each objReg In colSpedRegistro
    
        iX1 = iX1 + 1
    
        Call objTela.SPED_Status("Processando Registro 9900", CStr(iX1) & SEPARADOR & objECD.sTipo)
    
        objECD.lSeq = objECD.lSeq + 1
        
        objECD.sRegistro = "|9900|" & objReg.sTipo & "|" & objReg.lCount & "|"
        
        lTotal = lTotal + objReg.lCount
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
        If lErro <> SUCESSO Then gError 203222

    Next
    
    '************ registro 9990 *******************
    Call objTela.SPED_Status("Processando Registro 9990", "")
    
    objECD.sTipo = "9990"
    objECD.lSeq = objECD.lSeq + 1
    
    objECD.sRegistro = "|9990|" & colSpedRegistro.Count + 3 & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203223
    
    '************ registro 9999 *******************
    Call objTela.SPED_Status("Processando Registro 9999", "")
    
    objECD.sTipo = "9999"
    objECD.lSeq = objECD.lSeq + 1
    
    objECD.sRegistro = "|9999|" & lTotal & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
    If lErro <> SUCESSO Then gError 203224
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'I030'", _
    sRegistro, objECD.lNumIntSped, objECD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203225
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203226

    iPos = 1

    For iIndice = 1 To 4
        iPos = iPos + 1
        iPos = InStr(iPos, sRegistro, "|")
    Next
    
    objECD.sRegistro = Mid(sRegistro, 1, iPos) & lTotal & Mid(sRegistro, iPos + 1)
    
    lErro = Comando_Executar(alComando(17), "UPDATE SpedContabil SET Registro = ? WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'I030'", _
    objECD.sRegistro, objECD.lNumIntSped, objECD.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 203227
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'J900'", _
    sRegistro, objECD.lNumIntSped, objECD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203228
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203229

    iPos = 1

    For iIndice = 1 To 5
        iPos = iPos + 1
        iPos = InStr(iPos, sRegistro, "|")
    Next
    
    objECD.sRegistro = Mid(sRegistro, 1, iPos) & lTotal & Mid(sRegistro, iPos + 1)
    
    lErro = Comando_Executar(alComando(17), "UPDATE SpedContabil SET Registro = ? WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'J900'", _
    objECD.sRegistro, objECD.lNumIntSped, objECD.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 203230
    
    Set colSpedCont1 = Nothing
    Set colSpedContabil1 = Nothing
    Set colSpedContabil1_Lanc = Nothing
    Set colSpedContRes1 = Nothing
    Set colSpedContRes2 = Nothing
    Set colSpedRegistro = Nothing

    Call objTela.SPED_Status("Gerando arquivo", "")

    Open sDiretorio For Output As #1
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? ORDER BY Seq", _
    sRegistro, objECD.lNumIntSped, objECD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203231
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203232

    iX1 = 0
    Do While lErro = AD_SQL_SUCESSO
    
        iX1 = iX1 + 1
    
        Call objTela.SPED_Status("Gerando arquivo", "Linha: " & CStr(iX1))
    
        Print #1, sRegistro
    
        lErro = Comando_BuscarProximo(alComando(16))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203233
    
    Loop

    Close #1
    
    Call objTela.SPED_Status("Finalizando", "")

    '???? depois descomentar
'    lErro = Comando_Executar(alComando(18), "DELETE FROM SpedContabil WHERE  NumIntSped = ? AND FilialEmpresa = ?", _
'    objECD.lNumIntSped, objECD.iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 203234

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    'reativa os locks
    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 207914
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 203235
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_DIARIO_SPED_SUCESSO", sDiretorio)
    
    Gera_Sped_Contabil_Diario = SUCESSO
    
    Exit Function
    
Erro_Gera_Sped_Contabil_Diario:

    Gera_Sped_Contabil_Diario = gErr
    
    Call Gera_Sped_Contabil_Diario_Trata_Erro(gErr, iFilialEmpresa, dValorDocPos, dValorDocNeg, objCidades.sDescricao)
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    'reativa os locks
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
   
    Call Transacao_Rollback
   
    Exit Function
    
End Function

Private Sub Gera_Sped_Contabil_Diario_Trata_Erro(ByVal gErr As Long, ByVal iFilialEmpresa As Integer, ByVal dValorDocPos As Double, ByVal dValorDocNeg As Double, ByVal sCidade As String)

    Select Case gErr
    
        Case 203059, 203060, 203068, 203069, 203115, 203119, 203121, 203143, 203144, 203151, 203152, 203159, 203168, 203179, 203186 To 203188, 203194 To 203197, 203207, 203208, 203238, 203239, 203242, 203243, 203246, 203247, 203264 To 203268, 207328, 207329, 207330, 207335, 207338, 207350, 207352, 207353
        
        Case 203113
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 203114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203116, 203117
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 203118
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTAREFMODELO_NAO_CADASTRADO", gErr)

        Case 203120
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, iFilialEmpresa)

        Case 203122
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, sCidade)

        Case 203123, 203128, 203131, 203133, 203135, 203137, 203145, 203160, 203180, 203189, 203198, 203205, 203209, 203214, 203218, 203240, 203244, 203248
            Call Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_MAIOR_TAM_MAXIMO", gErr)

        Case 203124, 203125, 203129, 203132, 203134, 203136, 203138, 203146, 203153, 203154, 203157, 203161, 203162, 203165, 203169, 203181, 203183, 203190, 203199, 203206, 203210, 203212, 203213, 203215, 203219 To 203224, 203241, 203245, 203249, 204748, 207332, 207336, 207341
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDCONTABIL", gErr)

        Case 203126, 203127, 203130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDCADASTRO", gErr)

        Case 203139, 203140, 203172, 203184, 203185, 203191, 203254, 203255, 203260, 203261, 207333, 207334, 207337
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)

        Case 203141, 203142
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)

        Case 203147, 203148
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 203149, 203150, 203158, 203166, 203167, 203170, 203175, 203176, 203182, 203200, 203256, 203257
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", gErr)

        Case 203155, 203156, 203163, 203164, 203171, 207339, 207340, 207342, 207351
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE", gErr)

        Case 203173, 203174
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)

        Case 203177, 203178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)

        Case 203192, 203193, 203258, 203259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", gErr)

        Case 203201, 203202, 203203, 203204, 203211, 203262, 203263
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)

        Case 203216, 203217
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DNRC", gErr)

        Case 203225, 203226, 203228, 203229, 203231 To 203233
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDCONTABIL", gErr)

        Case 203227, 203230
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SPEDCONTABIL", gErr)

        Case 203234
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SPEDCONTABIL", gErr)

        Case 203235
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case 211569
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAOUTROS_OBRIGATORIA", gErr, Format(dValorDocPos, "STANDARD"), Format(dValorDocNeg, "STANDARD"))
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203236)
     
    End Select
    
    Exit Sub
    
End Sub

Private Function Periodo_Le_Sped(ByVal dtData As Date, ByVal objPeriodo As ClassPeriodo, ByVal lComando As Long) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodo

Dim lErro As Long
Dim tPeriodo As typePeriodo
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Periodo_Le_Sped

    iFilialEmpresa = giFilialEmpresa

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Periodo.Exercicio, Periodo.Periodo, NomeExterno, Lancamentos, Apurado, DataInicio, DataFim, Lote, Fechado FROM Periodo, PeriodosFilial WHERE FilialEmpresa = ? AND Periodo.DataInicio<=? AND Periodo.DataFim>=? AND Periodo.Exercicio = PeriodosFilial.Exercicio AND Periodo.Periodo = PeriodosFilial.Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, tPeriodo.iLote, tPeriodo.iFechado, iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 6047
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6048
    
    objPeriodo.iExercicio = tPeriodo.iExercicio
    objPeriodo.iPeriodo = tPeriodo.iPeriodo
    objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
    objPeriodo.iLancamentos = tPeriodo.iLancamentos
    objPeriodo.iApurado = tPeriodo.iApurado
    objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
    objPeriodo.dtDataFim = tPeriodo.dtDataFim
    objPeriodo.iLote = tPeriodo.iLote
    objPeriodo.iFechado = tPeriodo.iFechado
    
    Periodo_Le_Sped = SUCESSO

    Exit Function

Erro_Periodo_Le_Sped:

    Periodo_Le_Sped = Err
    
    Select Case Err
 
    Case 6047, 6048
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err)
        MsgBox ("filialempresa: " & CStr(iFilialEmpresa) & " " & CStr(dtData))
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153386)
        
    End Select
       
    Exit Function
    
End Function

Private Function ContaCcl_Saldo_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double, lComando1 As Long, lComando2 As Long) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim iIndice As Integer
Dim iExercicio As Integer
Dim dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer
Dim dValMvDia As Double

On Error GoTo Erro_ContaCcl_Saldo_Sped

    'percorrer registros gerando os saldos na data para as diversas contas
    dSaldo = 0
    iMes = Month(dtData)
    
    lErro = Comando_Executar(lComando1, "SELECT Exercicio, SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Ccl = ? AND Conta = ?  ORDER BY Exercicio", _
        iExercicio, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 203070
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203071
    
    If lErro = AD_SQL_SUCESSO Then dSaldo = dSldIni
    
    Do While lErro = AD_SQL_SUCESSO
        
            
        If iExercicio <> objPer.iExercicio Then
        
            dSaldo = Arredonda_Moeda(dSaldo - dDeb(1) + dCre(1) - dDeb(2) + dCre(2) - dDeb(3) + dCre(3) - dDeb(4) + dCre(4) - dDeb(5) + dCre(5) - dDeb(6) + dCre(6) - dDeb(7) + dCre(7) - dDeb(8) + dCre(8) - dDeb(9) + dCre(9) - dDeb(10) + dCre(10) - dDeb(11) + dCre(11) - dDeb(12) + dCre(12), 2)
            
        Else
        
            iIndice = 1
            Do While iIndice < iMes
                dSaldo = Arredonda_Moeda(dSaldo + dCre(iIndice) - dDeb(iIndice), 2)
                iIndice = iIndice + 1
            Loop
            
        End If
    
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203072
    
    Loop
    
    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(lComando2, "SELECT SUM(Cre-Deb) FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Conta = ? AND  Data >= ? AND Data < ? ", _
            dValMvDia, iFilialEmpresa, sCcl, sConta, objPer.dtDataInicio, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 203073
                
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203074
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Arredonda_Moeda(dSaldo + dValMvDia, 2)
        
    End If

    ContaCcl_Saldo_Sped = SUCESSO
     
    Exit Function
    
Erro_ContaCcl_Saldo_Sped:

    ContaCcl_Saldo_Sped = gErr
     
    Select Case gErr

        Case 203070 To 203072
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
        
        Case 203073 To 203074
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203075)
     
    End Select
   
    Exit Function

End Function

Private Function ContaCcl_Movimentacao_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double, lComando As Long) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_ContaCcl_Movimentacao_Sped

    dSaldoCre = 0
    dSaldoDeb = 0

    'obtem valor movimentado no periodo (mes) até a data
    lErro = Comando_Executar(lComando, "SELECT SUM(Cre) , SUM(Deb) FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Conta = ? AND Data >= ? AND Data <= ?  ", _
        dSaldoCre, dSaldoDeb, iFilialEmpresa, sCcl, sConta, dtDataIni, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 203081

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203082

    ContaCcl_Movimentacao_Sped = SUCESSO

    Exit Function

Erro_ContaCcl_Movimentacao_Sped:

    ContaCcl_Movimentacao_Sped = gErr

    Select Case gErr

        Case 203081, 203082
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203083)

    End Select

    Exit Function

End Function

Private Function Conta_Saldo_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double, lComando1 As Long, lComando2 As Long) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim iIndice As Integer
Dim iExercicio As Integer
Dim dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer
Dim dValMvDia As Double

On Error GoTo Erro_Conta_Saldo_Sped

    'percorrer registros gerando os saldos na data para as diversas contas
    dSaldo = 0
    iMes = Month(dtData)
    
    lErro = Comando_Executar(lComando1, "SELECT Exercicio, SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta = ? ORDER BY Exercicio", _
        iExercicio, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 203061
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203062
    
    If lErro = AD_SQL_SUCESSO Then dSaldo = dSldIni
    
    Do While lErro = AD_SQL_SUCESSO
        
            
        If iExercicio <> objPer.iExercicio Then
        
            dSaldo = Arredonda_Moeda(dSaldo - dDeb(1) + dCre(1) - dDeb(2) + dCre(2) - dDeb(3) + dCre(3) - dDeb(4) + dCre(4) - dDeb(5) + dCre(5) - dDeb(6) + dCre(6) - dDeb(7) + dCre(7) - dDeb(8) + dCre(8) - dDeb(9) + dCre(9) - dDeb(10) + dCre(10) - dDeb(11) + dCre(11) - dDeb(12) + dCre(12), 2)
            
        Else
        
            iIndice = 1
            Do While iIndice < iMes
                dSaldo = Arredonda_Moeda(dSaldo + dCre(iIndice) - dDeb(iIndice), 2)
                iIndice = iIndice + 1
            Loop
            
        End If
    
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203063
    
    Loop
    
    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(lComando2, "SELECT SUM(Cre-Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data >= ? AND Data < ? ", _
            dValMvDia, iFilialEmpresa, sConta, objPer.dtDataInicio, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 203064
                
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203065
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Arredonda_Moeda(dSaldo + dValMvDia, 2)
        
    End If
   
    Conta_Saldo_Sped = SUCESSO
     
    Exit Function
    
Erro_Conta_Saldo_Sped:

    Conta_Saldo_Sped = gErr
     
    Select Case gErr
          
        Case 203061 To 203063
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case 203064 To 203065
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203066)
     
    End Select
   
    Exit Function

End Function

Private Function Conta_Movimentacao_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double, lComando As Long) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Conta_Movimentacao_Sped

    dSaldoCre = 0
    dSaldoDeb = 0

    'obtem valor movimentado no periodo (mes) até a data
    lErro = Comando_Executar(lComando, "SELECT SUM(Cre), SUM(Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data >= ? AND Data <= ? ", _
        dSaldoCre, dSaldoDeb, iFilialEmpresa, sConta, dtDataIni, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 203077

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203078

    Conta_Movimentacao_Sped = SUCESSO

    Exit Function

Erro_Conta_Movimentacao_Sped:

    Conta_Movimentacao_Sped = gErr

    Select Case gErr

        Case 203077, 203078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203079)

    End Select

    Exit Function

End Function

Private Function SpedContabil1_Armazena_SldIni(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dSldIni As Double, ByVal objECD As ClassSpedContabil, ByVal dtDataIniAux As Date, ByVal dtDataFimAux As Date, ByVal dtDataIni As Date, ByVal sContaBD As String, ByVal iNaturezaSped As Integer, Optional ByVal iAfter As Integer = 0) As Long

'só os saldos iniciais do primeiro periodo que precisam ser armazenados
'os demais saldos sairão dos lançamentos

    Dim objECD1 As New ClassSpedContabil1
    Dim objECDAux As New ClassSpedContabil
    
    objECDAux.iFilialEmpresa = objECD.iFilialEmpresa
    objECDAux.lNumIntSped = objECD.lNumIntSped
    objECDAux.lSeq = objECD.lSeq
    objECDAux.sTipo = objECD.sTipo
    
    Set objECD1.objSpedContabil = objECDAux
    
    objECD1.sConta = sContaTela
    objECD1.sContaBD = sContaBD
    objECD1.sCcl = sCclMascarado
    If dtDataIni = dtDataIniAux Then objECD1.dSldIni = Arredonda_Moeda(dSldIni)
    objECD1.dtDataIni = dtDataIniAux
    objECD1.dtDataFim = dtDataFimAux
    objECD1.iNaturezaSped = iNaturezaSped
    
    
    If iAfter = 0 Then
        colSpedContabil1.Add objECD1
    Else
        colSpedContabil1.Add objECD1, , , iAfter
    End If
        
    SpedContabil1_Armazena_SldIni = SUCESSO

End Function

Private Function SpedContabil1_Atualiza_CredDeb(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dValor As Double, ByVal dtData As Date, ByVal iTipo As Integer) As Long

    Dim objECD1 As New ClassSpedContabil1
    Dim bAchou As Boolean
    
    bAchou = False
    For Each objECD1 In colSpedContabil1
    
        If objECD1.sConta = sContaTela And objECD1.sCcl = sCclMascarado And dtData >= objECD1.dtDataIni And dtData <= objECD1.dtDataFim Then
            If dValor > 0 Then
                objECD1.dCred = objECD1.dCred + dValor
            Else
                objECD1.dDeb = objECD1.dDeb - dValor
            End If
            bAchou = True
            Exit For
        
        End If
        
    Next
    
    If bAchou = False And iTipo = SPED_CONTAB_TIPO_NORMAL Then ' Se for o FCont vai ser normal não achar contas de naturezas que não devem ir no saldo
        MsgBox ("erro")
    End If
    
    SpedContabil1_Atualiza_CredDeb = SUCESSO

End Function

Private Function SpedContabil_FormulaDRE(ByVal objECD As ClassSpedContabil, ByVal iCodigo As Integer, lTotalRegBloco1 As Long, ByVal objRegI052 As ClassSpedRegistro, colSpedCont1 As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iCodigoDem As Integer
Dim iAchou As Integer
Dim objSpedCont1 As ClassSpedContabil1

On Error GoTo Erro_SpedContabil_FormulaDRE

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207343
    Next

    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM RelDREFormula WHERE Upper(Relatorio) = 'DRE' AND Upper(Modelo) = 'SPED' AND CodigoFormula = ?  ", _
    iCodigoDem, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 207344

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207345

    Do While lErro = AD_SQL_SUCESSO
        
        iAchou = 0
        For Each objSpedCont1 In colSpedCont1
            If objSpedCont1.sConta = iCodigoDem And objSpedCont1.sCcl = "" Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then
        
            Set objSpedCont1 = New ClassSpedContabil1
            
            objSpedCont1.sConta = iCodigoDem
            
            colSpedCont1.Add objSpedCont1

        
            '******** registro I052 *****************
'            objECD.sTipo = "I052"
'            objECD.lSeq = objECD.lSeq + 1
'            lTotalRegBloco1 = lTotalRegBloco1 + 1
'
'            objRegI052.lCount = objRegI052.lCount + 1
'
'            objECD.sRegistro = "|I052||" & iCodigoDem & "|"
'
'            lErro = Comando_Executar(alComando(2), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'            objECD.lNumIntSped, objECD.lSeq, objECD.iFilialEmpresa, objECD.sTipo, objECD.sRegistro)
'            If lErro <> SUCESSO Then gError 207346

        End If

        lErro = SpedContabil_FormulaDRE(objECD, iCodigoDem, lTotalRegBloco1, objRegI052, colSpedCont1)
        If lErro <> SUCESSO Then gError 207347

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207348

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    SpedContabil_FormulaDRE = SUCESSO
     
    Exit Function
    
Erro_SpedContabil_FormulaDRE:

    SpedContabil_FormulaDRE = gErr
     
    Select Case gErr
          
        Case 207343
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207346
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDCONTABIL", gErr)

        Case 207344, 207345, 207348
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDREFORMULA", gErr)
            
        Case 207347

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207349)
     
    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
                            
Private Function SpedContabil1_SaldoFinal_SaldoInicial(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dValor As Double, ByVal dtData As Date, iAchou As Integer) As Long
'colocar o saldo final como o saldo inicial do proximo periodo

Dim objECD1 As New ClassSpedContabil1
Dim dtDataProxPer As Date
Dim dtDataFimProxPer As Date
Dim dtDataAux As Date
Dim objECD As New ClassSpedContabil
Dim objECD1_Aux As ClassSpedContabil1
Dim dtDataAnt As Date
Dim iContador As Integer
    
On Error GoTo Erro_SpedContabil1_SaldoFinal_SaldoInicial
    
    iAchou = 0
    
    'se tem um valor a ser passado de um mes para o outro
    dtDataAnt = dtData
    dtDataAux = DateAdd("m", 1, dtData)
    dtDataProxPer = CDate("01/" & Month(dtDataAux) & "/" & Year(dtDataAux))
    
    'pesquisa se tem no periodo imediatamente contiguo um registro para colocar o saldo
    For Each objECD1 In colSpedContabil1
   
'       If objECD1.sConta = sContaTela And objECD1.sCcl = sCclMascarado And dtDataProxPer = objECD1.dtDataIni Then
        If objECD1.sConta = sContaTela And objECD1.sCcl = sCclMascarado And objECD1.dtDataIni = dtDataProxPer Then

            objECD1.dSldIni = Arredonda_Moeda(dValor)
            iAchou = 1
           
            Exit For
   
        End If
       
    Next
    
    SpedContabil1_SaldoFinal_SaldoInicial = SUCESSO
    
    Exit Function
    
Erro_SpedContabil1_SaldoFinal_SaldoInicial:

    SpedContabil1_SaldoFinal_SaldoInicial = gErr
     
    Select Case gErr
          
        Case 207361
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207362)
     
    End Select
   
    Exit Function

End Function

Private Function SpedContabil1_SaldoInicial(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dDiferenca As Double, ByVal dtData As Date, ByVal dtDataFim As Date, ByVal sContaBD As String, ByRef lTotalRegBloco1 As Long, ByVal objRegI155 As ClassSpedRegistro, ByVal iNaturezaSped As Integer) As Long
'colocar o saldo final como o saldo inicial do proximo periodo

Dim objECD1 As New ClassSpedContabil1
Dim dtDataProxPer As Date
Dim dtDataFimProxPer As Date
Dim dtDataAux As Date
Dim iAchou As Integer
Dim objECD As New ClassSpedContabil
Dim objECD1_Aux As ClassSpedContabil1
Dim objECD1_Prox As ClassSpedContabil1
Dim dtDataAnt As Date
Dim iContador As Integer
Dim iTotal As Integer
Dim lErro As Long
    
On Error GoTo Erro_SpedContabil1_SaldoInicial
    
    'se tem um valor a ser passado de um mes para o outro
    dtDataAnt = dtData
    dtDataAux = DateAdd("m", 1, dtData)
    dtDataProxPer = CDate("01/" & Month(dtDataAux) & "/" & Year(dtDataAux))
    dtDataFimProxPer = DateAdd("m", 1, dtDataProxPer) - 1
    
    Do While dtDataProxPer <= dtDataFim
    
        If dtDataFimProxPer > dtDataFim Then dtDataFimProxPer = dtDataFim
        
        iAchou = 0
       
        'pesquisa se tem no periodo imediatamente contiguo um registro para colocar o saldo
        For Each objECD1 In colSpedContabil1
       
       
    '       If objECD1.sConta = sContaTela And objECD1.sCcl = sCclMascarado And dtDataProxPer = objECD1.dtDataIni Then
            If objECD1.sConta = sContaTela And objECD1.sCcl = sCclMascarado And objECD1.dtDataIni = dtDataProxPer Then
    
                iAchou = 1
                objECD1.dSldIni = objECD1.dSldIni - Arredonda_Moeda(dDiferenca)
               
                Exit For
       
            End If
           
        Next
        
        'se nao achou o I155 do proximo mes, significa que o saldo final foi criado devido
        'a diferenca entre credito e debito inicial ==> cria um I155 agora para todos os meses em sequencia e coloca o saldo incial
        'tem q verificar se ja nao podia existir o registro
        If iAchou = 0 Then
        
            iTotal = colSpedContabil1.Count
        
            For iContador = 2 To iTotal
        
                Set objECD1_Prox = colSpedContabil1(iContador)
                'se encontrou o primeiro I155 depois do mes que esta sendo o mes em questao
                'o sequencial vai ser o sequencial dele - 1
                If objECD1_Prox.dtDataIni > dtDataProxPer Then
        
                    
                    Set objECD1_Aux = colSpedContabil1(iContador - 1)
                    
                    Set objECD = New ClassSpedContabil
        
                    objECD.iFilialEmpresa = objECD1_Prox.objSpedContabil.iFilialEmpresa
                    objECD.lNumIntSped = objECD1_Prox.objSpedContabil.lNumIntSped
                    objECD.sTipo = objECD1_Prox.objSpedContabil.sTipo
                    objECD.lSeq = objECD1_Aux.objSpedContabil.lSeq + 1
            
                    objRegI155.lCount = objRegI155.lCount + 1
                    lTotalRegBloco1 = lTotalRegBloco1 + 1
            
                    lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, sCclMascarado, -dDiferenca, objECD, dtDataProxPer, dtDataFimProxPer, dtDataProxPer, sContaBD, iNaturezaSped, iContador - 1)
                    If lErro <> SUCESSO Then gError 207361
                    
                    
                    Exit For
                End If
                
            Next
        
        End If
            
        dtDataAnt = dtDataProxPer
        dtDataProxPer = DateAdd("m", 1, dtDataProxPer)
        dtDataFimProxPer = DateAdd("m", 1, dtDataProxPer) - 1
        
            
    Loop
    
    SpedContabil1_SaldoInicial = SUCESSO
    
    Exit Function
    
Erro_SpedContabil1_SaldoInicial:

    SpedContabil1_SaldoInicial = gErr
     
    Select Case gErr
          
        Case 207361
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207362)
     
    End Select
   
    Exit Function

End Function

Function SFF_Texto(ByVal sTexto As String, Optional ByVal lTamMax As Long = 0, Optional bFazTrim As Boolean = True, Optional bRemoveEspacoDuplo As Boolean = True, Optional bRemoveAcentos As Boolean = False, Optional bUpperCase As Boolean = False) As String
Dim sTextoAux As String
Dim iIndice As Integer
Dim iPos As Integer
'1 Formato dos campos
'ALFANUMÉRICO(C): representados por “C” - todos os caracteres, excetuados os caracteres
'“|”(Pipe ou Barra Vertical: caractere 124 da Tabela ASCII) e os não-imprimíveis (caracteres 00 a 31 da Tabela ASCII);
'NUMÉRICO(N): representados por “N” - algarismos das posições 'de 48 a 57 da Tabela ASCII.

    sTextoAux = sTexto
    
    sTextoAux = Replace(sTextoAux, Chr(124), "") 'Remove pipe "|"
    
    For iIndice = 0 To 31 'Remove caracteres não-imprimíveis
        sTextoAux = Replace(sTextoAux, Chr(iIndice), "")
    Next
    
    If bFazTrim Then 'Remove espaços no início e no final (OPCIONAL)
        sTextoAux = Trim(sTextoAux)
    End If
    
    If bRemoveEspacoDuplo Then 'Remove espaços duplos "  " (OPCIONAL)
        iPos = InStr(1, sTextoAux, "  ")
        Do While iPos <> 0
            sTextoAux = Replace(sTextoAux, "  ", " ")
            iPos = InStr(1, sTextoAux, "  ")
        Loop
    End If
    
    If bRemoveAcentos Then 'Tira acentuação (OPCIONAL)
        sTextoAux = DesacentuaTexto(sTextoAux)
    End If
    
    If bUpperCase Then 'Muda para maiúsculo (OPCIONAL)
        sTextoAux = UCase(sTextoAux)
    End If
    
    If lTamMax > 0 Then 'Reduz o tamanho do texto (OPCIONAL)
        sTextoAux = left(sTextoAux, lTamMax)
    End If
    
    SFF_Texto = sTextoAux

    'SFF_Texto = left(Replace(Replace(Replace(sTexto, vbTab, " "), Chr(13), ""), Chr(10), ""), lTamMax)
End Function

Function SFF_CD(ByVal dValor As Double) As String
Dim sTextoAux As String
Dim dValorAux As Double

    dValorAux = StrParaDbl(Format(dValor, "fixed"))
    
    If dValorAux > DELTA_VALORMONETARIO Then
        sTextoAux = "C"
    ElseIf dValorAux < -DELTA_VALORMONETARIO Then
        sTextoAux = "D"
    Else
        sTextoAux = "C"
    End If
    
    SFF_CD = sTextoAux

End Function

