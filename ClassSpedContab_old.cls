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

Function Gera_Sped_Contabil_Diario(ByVal sDiretorio As String, ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lNumOrd As Long, ByVal objTela As Object) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim lNumIntSped As Long
Dim objSpedContabil As ClassSpedContabil
Dim lSeq As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEndereco As New ClassEndereco
Dim objCidades As New ClassCidades
Dim lTransacao As Long
Dim alComando(1 To 30) As Long
Dim sCodigoInst As String
Dim sCodigoCad As String
Dim tPlanoConta As typePlanoConta
Dim iNaturezaSped As Integer
Dim iNatSpedAtual As Integer
Dim sContaTela As String
Dim sContaPai As String
Dim lModelo As Long
Dim sContaRef As String
Dim sCclIni As String
Dim sCclFim As String
Dim sContaTelaRef As String
Dim sCclIniDem As String
Dim sCclFimDem As String
Dim iCodigoDem As Integer
Dim sCcl As String
Dim sCclMascarado As String
Dim iUsoCcl As Integer
Dim sDescCcl As String
Dim dSaldoIni As Double
Dim dSaldoDeb As Double
Dim dSaldoCre As Double
Dim dSaldoFim As Double
Dim sConta As String
Dim tLanc As typeLancamento
Dim sOrigemAtual As String
Dim iExercicioAtual As Integer
Dim iPeriodoLanAtual As Integer
Dim lDocAtual As Long
Dim iFilialEmpresaAtual As Integer
Dim sIdentificador As String
Dim lTotalRegBloco1 As Long
Dim sDescQualiSigCTB As String
Dim objSpedRegistro As ClassSpedRegistro
Dim colSpedRegistro As New Collection
Dim objSpedRegistroI050 As ClassSpedRegistro
Dim objSpedRegistroI051 As ClassSpedRegistro
Dim objSpedRegistroI052 As ClassSpedRegistro
Dim objSpedRegistroI250 As ClassSpedRegistro
Dim sRegistro As String
Dim iPos As Integer
Dim dtData As Date
Dim sCclMascarada As String
Dim iExerAnt As Integer
Dim objPer As New ClassPeriodo
Dim sContaAtual As String
Dim dSaldoAcum As Double
Dim dSaldoCreAcum As Double
Dim dSaldoDebAcum As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lContador As Long
Dim dSaldoIniTotal As Double
Dim dSaldoFimTotal As Double
Dim sContaTelaUltimo As String
Dim dSaldoIniUltimo As Double
Dim dSaldoFimUltimo As Double
Dim lSeqUltimo As Long
Dim dSaldoCreUltimo As Double
Dim dSaldoDebUltimo As Double
Dim iX1 As Double, iX2 As Double, iX3 As Double
Dim lSeqDoc As Long
Dim dtDataAtual As Date
Dim dValorDocPos As Double
Dim dValorDocNeg As Double
Dim objLanc As ClassLancamento_Detalhe
Dim colLanc As New Collection
Dim colSpedContabil1 As New Collection
Dim objSpedContabil1 As ClassSpedContabil1
Dim dValorDoc As Double

On Error GoTo Erro_Gera_Sped_Contabil_Diario


    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 203113

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
    
    
    
    If iUsoCcl = 2 Then
    
        Call objTela.SPED_Status("Totalizando Registros", "Ccl")
    
        lErro = Comando_Executar(alComando(21), "SELECT Count(*) From Ccl", lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 203256
        
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203257
        
        objTela.dValorTotal = objTela.dValorTotal + lContador
    
        Call objTela.SPED_Status("Totalizando Registros", "Conta\Ccl")
    
'        lErro = Comando_Executar(alComando(21), "SELECT COUNT( DISTINCT Conta+CCl) From MvPerCcl", lContador)
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
    
    Set objSpedContabil = New ClassSpedContabil
    
    ' descobre o plano de contas referencial que sera usado
    lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM PlanoContaRefModelo WHERE Oficial = 1 AND AnoVigencia <= ? ORDER BY AnoVigencia DESC", _
    lModelo, Year(dtDataFim))
    If lErro <> AD_SQL_SUCESSO Then gError 203116

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203117
    
    If lErro <> AD_SQL_SUCESSO Then gError 203118
    
    
    '******** registro 0000 *****************
    
    Call objTela.SPED_Status("Processando Registro 0000", "")
    
    objSpedContabil.sRegistro = "|0000|LECD|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    If objFilialEmpresa.iCodFilial = EMPRESA_TODA Then objFilialEmpresa.iCodFilial = 1
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 203119

    If lErro <> SUCESSO Then gError 203120

    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.sNome & "|" & objFilialEmpresa.sCgc & "|" & objFilialEmpresa.objEndereco.sSiglaEstado & "|"
    
    objCidades.sDescricao = objFilialEmpresa.objEndereco.sCidade
    
    lErro = CF("Cidade_Le_Nome", objCidades)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203121
    
    If lErro = ERRO_OBJETO_NAO_CADASTRADO Then gError 203122
    
    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.sInscricaoEstadual & "|" & objCidades.sCodIBGE & "|" & objFilialEmpresa.sInscricaoMunicipal & "||"
    
    objSpedContabil.sTipo = "0000"
    objSpedContabil.iFilialEmpresa = iFilialEmpresa
    objSpedContabil.lNumIntSped = lNumIntSped
    objSpedContabil.lSeq = 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203123
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203124

    '******** registro 0001 *****************

    Call objTela.SPED_Status("Processando Registro 0001", "")

    objSpedContabil.sTipo = "0001"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    objSpedContabil.sRegistro = "|0001|0|"
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203125

    '******** registro 0007 *****************

    Call objTela.SPED_Status("Processando Registro 0007", "")

    objSpedContabil.sTipo = "0007"
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo

    colSpedRegistro.Add objSpedRegistro
    
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
    
        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
        Call objTela.SPED_Status("Processando Registro 0007", CStr(objSpedContabil.lSeq) & SEPARADOR & sCodigoInst)
    
        objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
        objSpedContabil.sRegistro = "|0007|" & sCodigoInst & "|" & sCodigoCad & "|"
    
        If Len(objSpedContabil.sRegistro) > 500 Then gError 203128
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
        If lErro <> SUCESSO Then gError 203129
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203130
        
    Loop

    '******** registro 0990 *****************

    Call objTela.SPED_Status("Processando Registro 0990", "")

    objSpedContabil.sTipo = "0990"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedContabil.sRegistro = "|0990|" & CStr(objSpedContabil.lSeq) & "|"
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203131
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203132


    '******** registro I001 *****************

    Call objTela.SPED_Status("Processando Registro I001", "")

    objSpedContabil.sTipo = "I001"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedContabil.sRegistro = "|I001|0|"
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203133
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203134

    '******** registro I010 *****************
    
    Call objTela.SPED_Status("Processando Registro I010", "")

    objSpedContabil.sTipo = "I010"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    
    objSpedContabil.sRegistro = "|I010|G|1.00|"
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203135
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203136

    '******** registro I030 *****************

    Call objTela.SPED_Status("Processando Registro I030", "")

    objSpedContabil.sTipo = "I030"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    
    objSpedContabil.sRegistro = "|I030|TERMO DE ABERTURA|" & lNumOrd & "|Escrituração Contábil Digital do Livro Diário Geral||"
    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.sNome & "|" & objFilialEmpresa.sJucerja & "|"
    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.sCgc & "|" & Format(objFilialEmpresa.dtDataJucerja, "ddmmyyyy") & "||"
    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.objEndereco.sCidade & "|"
    
    
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203137
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203138


    '******** registro I050 *****************

    Call objTela.SPED_Status("Processando Registro I050", "")

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT Conta, TipoConta, NivelConta, DescConta, NaturezaSped FROM PlanoConta ORDER BY COnta", _
    tPlanoConta.sConta, tPlanoConta.iTipoConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, iNaturezaSped)
    If lErro <> AD_SQL_SUCESSO Then gError 203139

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203140

    Set objSpedRegistroI050 = New ClassSpedRegistro
    colSpedRegistro.Add objSpedRegistroI050
    objSpedRegistroI050.sTipo = "I050"
    
    Set objSpedRegistroI051 = New ClassSpedRegistro
    colSpedRegistro.Add objSpedRegistroI051
    objSpedRegistroI051.sTipo = "I051"

    Set objSpedRegistroI052 = New ClassSpedRegistro
    colSpedRegistro.Add objSpedRegistroI052
    objSpedRegistroI052.sTipo = "I052"

    Do While lErro = AD_SQL_SUCESSO
    
        Call objTela.SPED_Status("Processando Registro I050", CStr(objSpedContabil.lSeq + 1) & SEPARADOR & tPlanoConta.sConta)
    
        If tPlanoConta.iNivelConta = 1 Then iNatSpedAtual = iNaturezaSped
    
        lErro = Comando_Executar(alComando(4), "SELECT DataAtualizacao FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
            dtData, tPlanoConta.sConta, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203141
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203142
    
        If lErro = AD_SQL_SEM_DADOS Then dtData = StrParaDate("01/01/2009")
    
        objSpedContabil.sTipo = "I050"
        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
    
        objSpedRegistroI050.lCount = objSpedRegistroI050.lCount + 1

    
        objSpedContabil.sRegistro = "|I050|" & Format(dtData, "ddmmyyyy") & "|" & Format(iNatSpedAtual, "00") & "|" & IIf(tPlanoConta.iTipoConta = 1, "S", "A") & "|"
        objSpedContabil.sRegistro = objSpedContabil.sRegistro & tPlanoConta.iNivelConta & "|"
        
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
        
        objSpedContabil.sRegistro = objSpedContabil.sRegistro & sContaTela & "|" & sContaPai & "|" & tPlanoConta.sDescConta & "|"
    
        If Len(objSpedContabil.sRegistro) > 500 Then gError 203145
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
        If lErro <> SUCESSO Then gError 203146
        
        objTela.dValorAtual = objTela.dValorAtual + 1
        objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
    
        objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
    
        DoEvents
        
        If objTela.iCancelaBatch = CANCELA_BATCH Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
        
            If vbMsgRes = vbYes Then gError 203265
            
            objTela.iCancelaBatch = 0
                
        End If
        
        
        'se for uma conta analitica
        If tPlanoConta.iTipoConta = 3 Then
        
            Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objSpedContabil.lSeq) & SEPARADOR & tPlanoConta.sConta)
        
            sContaRef = String(STRING_CONTA, 0)
            sContaTelaRef = String(STRING_CONTA, 0)
            sCclIni = String(STRING_CCL, 0)
            sCclFim = String(STRING_CCL, 0)
        
            lErro = Comando_Executar(alComando(6), "SELECT P.ContaImp, PC.Conta, PC.CclInicial, PC.CclFinal FROM PlanoContaRefConfig AS PC, PlanoContaRef As P WHERE PC.CodigoModelo = P.CodigoModelo AND PC.Conta = P.Conta AND Tipo = 3 AND PC.CodigoModelo = ? AND ContaInicial <= ? And ContaFinal >= ? ", _
            sContaTelaRef, sContaRef, sCclIni, sCclFim, lModelo, tPlanoConta.sConta, tPlanoConta.sConta)
            If lErro <> AD_SQL_SUCESSO Then gError 203147
        
            lErro = Comando_BuscarPrimeiro(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203148
        
            iX1 = 0
            Do While lErro = AD_SQL_SUCESSO
            
                iX1 = iX1 + 1
            
                Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objSpedContabil.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sContaTelaRef)
            
                If sCclIni <> "" And sCclFim <> "" Then
                    'se o centro de custo da conta  estiver preenchido
                
                    sCcl = String(STRING_CCL, 0)
                    
                    lErro = Comando_Executar(alComando(8), "SELECT Ccl FROM Ccl WHERE TipoCcl = 3 And Ccl >= ? AND Ccl <= ?", _
                    sCcl, sCclIni, sCclFim)
                    If lErro <> AD_SQL_SUCESSO Then gError 203149
                    
                    lErro = Comando_BuscarPrimeiro(alComando(8))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                    
                    iX2 = 0
                    Do While lErro = AD_SQL_SUCESSO
                        
                        iX2 = iX2 + 1
                        
                        Call objTela.SPED_Status("Processando Registro I050 - Referencial", CStr(objSpedContabil.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sContaTelaRef & SEPARADOR & sCcl)
                        
                        lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                        If lErro <> SUCESSO Then gError 203151
                                    
'                        lErro = Mascara_Mascarar_ContaRef(sContaRef, sContaTelaRef)
'                        If lErro <> SUCESSO Then gError 203152
                        
                        'sContaTelaRef = sContaTelaRef & ".00"
                
                        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                        objSpedContabil.sRegistro = "|I051|10|" & sCclMascarado & "|" & sContaTelaRef & "|"
                        lTotalRegBloco1 = lTotalRegBloco1 + 1
                    
                        objSpedRegistroI051.lCount = objSpedRegistroI051.lCount + 1
                
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                        If lErro <> SUCESSO Then gError 203153
                    
                        'se for conta de ativo ou passivo a conta aglutinadora é a propria conta me questao
                        If iNatSpedAtual = 1 Or iNatSpedAtual = 2 Then
                            
                            '******** registro I052 *****************
                            objSpedContabil.sTipo = "I052"
                            objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                            lTotalRegBloco1 = lTotalRegBloco1 + 1
                        
                            objSpedRegistroI052.lCount = objSpedRegistroI052.lCount + 1
                        
                            objSpedContabil.sRegistro = "|I052||" & sContaTela & "|"
                        
                            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                            objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                            If lErro <> SUCESSO Then gError 203154
                            
                        Else
                    
                            'se for conta de demonstrativo ==> pegar o demonstrativo SPED para a conta em questao
                        
                            lErro = Comando_Executar(alComando(7), "SELECT R.Codigo FROM RelDRE As R, RelDREConta As RC WHERE R.Relatorio = RC.Relatorio AND R.Modelo = RC.Modelo AND R.Codigo = RC.Codigo AND Tipo = 0 AND Upper(RC.Relatorio) = 'DRE' AND Upper(RC.Modelo) = 'SPED' AND ContaInicial <= ? And ContaFinal >= ? AND CclInicial <= ? AND CclFinal >=? ", _
                            iCodigoDem, tPlanoConta.sConta, tPlanoConta.sConta, sCcl, sCcl)
                            If lErro <> AD_SQL_SUCESSO Then gError 203155
                    
                            lErro = Comando_BuscarPrimeiro(alComando(7))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203156
                    
                            If lErro = AD_SQL_SUCESSO Then
                                
                                '******** registro I052 *****************
                                objSpedContabil.sTipo = "I052"
                                objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                                lTotalRegBloco1 = lTotalRegBloco1 + 1
                            
                                objSpedRegistroI052.lCount = objSpedRegistroI052.lCount + 1
                            
                                objSpedContabil.sRegistro = "|I052|" & sCclMascarado & "|" & iCodigoDem & "|"
                            
                                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                                objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                                If lErro <> SUCESSO Then gError 203157
                        
                            End If
                        
                        End If
                        
                        
                        lErro = Comando_BuscarProximo(alComando(8))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203158
                    
                    Loop
                
                Else
                
                    'se o centro de custo da conta estiver zerado
                
'                    lErro = Mascara_Mascarar_ContaRef(sContaRef, sContaTelaRef)
'                    If lErro <> SUCESSO Then gError 203159
                    
                '    sContaTelaRef = sContaTelaRef & ".00"
                    
                    objSpedContabil.sRegistro = "|I051|10||" & sContaTelaRef & "|"
                
                    If Len(objSpedContabil.sRegistro) > 500 Then gError 203160
                
                
                    '******** registro I051 *****************
                    objSpedContabil.sTipo = "I051"
                    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                    lTotalRegBloco1 = lTotalRegBloco1 + 1
            
                    objSpedRegistroI051.lCount = objSpedRegistroI051.lCount + 1
            
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                    If lErro <> SUCESSO Then gError 203161
                    
                                            
                    'se for conta de ativo ou passivo a conta aglutinadora é a propria conta me questao
                    If iNatSpedAtual = 1 Or iNatSpedAtual = 2 Then
                        
                        '******** registro I052 *****************
                        objSpedContabil.sTipo = "I052"
                        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                        lTotalRegBloco1 = lTotalRegBloco1 + 1
    
                        objSpedRegistroI052.lCount = objSpedRegistroI052.lCount + 1
                    
                        objSpedContabil.sRegistro = "|I052||" & sContaTela & "|"
                    
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                        If lErro <> SUCESSO Then gError 203162
                        
                    Else
                    
                        Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objSpedContabil.lSeq) & SEPARADOR & tPlanoConta.sConta)
                    
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
                    
                            Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objSpedContabil.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & CStr(iCodigoDem))
                    
                            If sCclIniDem = "" And sCclFimDem = "" Then
                            
                                '******** registro I052 *****************
                                objSpedContabil.sTipo = "I052"
                                objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                                lTotalRegBloco1 = lTotalRegBloco1 + 1
                            
                                objSpedRegistroI052.lCount = objSpedRegistroI052.lCount + 1
                            
                                objSpedContabil.sRegistro = "|I052||" & iCodigoDem & "|"
                            
                                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                                objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                                If lErro <> SUCESSO Then gError 203165
                                
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
                                        
                                        Call objTela.SPED_Status("Processando Registro I050 - DRE", CStr(objSpedContabil.lSeq) & SEPARADOR & CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & CStr(iCodigoDem) & SEPARADOR & sCcl)
                                       
                                        lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                                        If lErro <> SUCESSO Then gError 203168
                                                    
                                        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                                        objSpedContabil.sRegistro = "|I052|" & sCclMascarado & "|" & iCodigoDem & "|"
                                        lTotalRegBloco1 = lTotalRegBloco1 + 1
        
                                        objSpedRegistroI052.lCount = objSpedRegistroI052.lCount + 1
                                        
                                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                                        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                                        If lErro <> SUCESSO Then gError 203169
                                    
                                        lErro = Comando_BuscarProximo(alComando(8))
                                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203170
                                    
                                    Loop
                                
                                End If
                            
                            End If
                            
                            lErro = Comando_BuscarProximo(alComando(7))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203171
                            
                        Loop
                    
                    End If
            
                End If
            
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203148
            
            Loop
        
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
    
        lErro = Comando_Executar(alComando(10), "SELECT Ccl, DescCcl FROM Ccl ORDER BY Ccl", sCcl, sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 203175
    
        lErro = Comando_BuscarPrimeiro(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203176
    
        Set objSpedRegistro = New ClassSpedRegistro
        
        colSpedRegistro.Add objSpedRegistro
    
        objSpedRegistro.sTipo = "I100"
    
        iX1 = 0
        Do While lErro = AD_SQL_SUCESSO
        
            iX1 = iX1 + 1
        
            Call objTela.SPED_Status("Processando Registro I100 - Ccl", CStr(iX1) & SEPARADOR & sCcl)
        
            lErro = Comando_Executar(alComando(11), "SELECT DataAtualizacao FROM CclHistorico WHERE CodigoCcl = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                dtData, sCcl, dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 203177
        
            lErro = Comando_BuscarPrimeiro(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203178
        
            If lErro <> AD_SQL_SUCESSO Then dtData = StrParaDate("01/01/2009")
        
            objSpedContabil.sTipo = "I100"
            objSpedContabil.lSeq = objSpedContabil.lSeq + 1
            lTotalRegBloco1 = lTotalRegBloco1 + 1

            objSpedRegistro.lCount = objSpedRegistro.lCount + 1

            objSpedContabil.sRegistro = "|I100|" & Format(dtData, "ddmmyyyy") & "|"
            
            lErro = Mascara_MascararCcl(sCcl, sCclMascarada)
            If lErro <> SUCESSO Then gError 203179
            
            objSpedContabil.sRegistro = objSpedContabil.sRegistro & sCclMascarada & "|" & sDescCcl & "|"
        
            If Len(objSpedContabil.sRegistro) > 500 Then gError 203180
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
            If lErro <> SUCESSO Then gError 203181

            objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
            objTela.dValorAtual = objTela.dValorAtual + 1
        
            objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
        
            DoEvents
            
            If objTela.iCancelaBatch = CANCELA_BATCH Then
        
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
            
                If vbMsgRes = vbYes Then gError 203266
                
                objTela.iCancelaBatch = 0
                    
            End If

            lErro = Comando_BuscarProximo(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203182
            
        Loop

    End If


    '******** registro I150 *****************
    
    Call objTela.SPED_Status("Processando Registro I150", "")
    
    objSpedContabil.sTipo = "I150"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    objSpedContabil.sRegistro = "|I150|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"

    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203183


    '******** registro I155 *****************

    Call objTela.SPED_Status("Processando Registro I155", "")

    objSpedContabil.sTipo = "I155"
    
    tPlanoConta.sConta = String(STRING_CONTA, 0)
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    
    'se nao usa ccusto extra-contabil
    If iUsoCcl <> 2 Then
    
        lErro = CF("Periodo_Le", dtDataIni, objPer)
        If lErro <> SUCESSO Then gError 203059
    
        Call objTela.SPED_Status("Processando Registro I155", "Calcula os Saldos")
    
        'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
        lErro = CF("RelCtaSaldo_Prepara1", objPer, alComando(20), iExerAnt)
        If lErro <> SUCESSO Then gError 203060
    
        lErro = Comando_Executar(alComando(12), "SELECT Conta FROM PlanoConta WHERE TipoConta = 3 ORDER BY COnta", _
        tPlanoConta.sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 203184
    
        lErro = Comando_BuscarPrimeiro(alComando(12))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203185
    
        iX1 = 0
        Do While lErro = AD_SQL_SUCESSO
        
            iX1 = iX1 + 1
        
            Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & tPlanoConta.sConta)
        
            'lErro = CF("Conta_Saldo", iFilialEmpresa, tPlanoConta.sConta, dtDataIni, iExerAnt, objPer, dSaldoIni)
            lErro = Conta_Saldo_Sped(iFilialEmpresa, tPlanoConta.sConta, dtDataIni, iExerAnt, objPer, dSaldoIni, alComando(27), alComando(28))
            If lErro <> SUCESSO Then gError 203186
            
            'lErro = CF("Conta_Movimentacao", iFilialEmpresa, tPlanoConta.sConta, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb)
            lErro = Conta_Movimentacao_Sped(iFilialEmpresa, tPlanoConta.sConta, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb, alComando(29))
            If lErro <> SUCESSO Then gError 203187
            
            If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
        
                objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                lTotalRegBloco1 = lTotalRegBloco1 + 1
                
                lErro = Mascara_MascararConta(tPlanoConta.sConta, sContaTela)
                If lErro <> SUCESSO Then gError 203188
                
                dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
                
                objSpedContabil.sRegistro = "|I155|" & sContaTela & "||" & Format(Abs(dSaldoIni), "fixed") & "|" & IIf(dSaldoIni > 0, "C", "D") & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & IIf(dSaldoFim > 0, "C", "D") & "|"
                
                objSpedRegistro.lCount = objSpedRegistro.lCount + 1
                
                lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, "", dSaldoIni, objSpedContabil)
                If lErro <> SUCESSO Then gError 207329
                
                
'                If Len(objSpedContabil.sRegistro) > 500 Then gError 203189
'
'                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'                objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
'                If lErro <> SUCESSO Then gError 203190
        
                objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
                objTela.dValorAtual = objTela.dValorAtual + 1
            
                objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
            
                DoEvents
                
                If objTela.iCancelaBatch = CANCELA_BATCH Then
            
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
                
                    If vbMsgRes = vbYes Then gError 203267
                    
                    objTela.iCancelaBatch = 0
                        
                End If
        
        
        
            End If
        
            lErro = Comando_BuscarProximo(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203191
        
        
        Loop
        
    'se usa ccusto extra contabil
    Else
    
        'lErro = CF("Periodo_Le", dtDataIni, objPer)
        lErro = Periodo_Le_Sped(dtDataIni, objPer, alComando(24))
        If lErro <> SUCESSO Then gError 203068
        
        Call objTela.SPED_Status("Processando Registro I155", "Calcula os Saldos")
       
        'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
        lErro = CF("RelCtaSaldo_Prepara1", objPer, alComando(20), iExerAnt)
        If lErro <> SUCESSO Then gError 203069
    
        dSaldoIniTotal = 0
        dSaldoFimTotal = 0
    
        tPlanoConta.sConta = String(STRING_CONTA, 0)
    
        lErro = Comando_Executar(alComando(22), "SELECT Conta FROM PlanoConta WHERE TipoConta = 3 ORDER BY COnta", _
        tPlanoConta.sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 204745
    
        lErro = Comando_BuscarPrimeiro(alComando(22))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204746
    
        sContaAtual = tPlanoConta.sConta
        dSaldoAcum = 0
        dSaldoCreAcum = 0
        dSaldoDebAcum = 0
            
        iX1 = 0
        Do While lErro = AD_SQL_SUCESSO
        
            iX1 = iX1 + 1
    
            Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & tPlanoConta.sConta)
   
            sConta = tPlanoConta.sConta
            sCcl = String(STRING_CCL, 0)
    
            lErro = Comando_Executar(alComando(13), "SELECT Ccl FROM MvPerCcl WHERE Conta = ? GROUP BY Ccl ORDER BY Ccl", _
            sCcl, tPlanoConta.sConta)
            If lErro <> AD_SQL_SUCESSO Then gError 203192
        
            lErro = Comando_BuscarPrimeiro(alComando(13))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203193
        
            iX2 = 0
            Do While lErro = AD_SQL_SUCESSO
            
                iX2 = iX2 + 1
        
                Call objTela.SPED_Status("Processando Registro I155", CStr(iX1) & SEPARADOR & CStr(iX2) & SEPARADOR & tPlanoConta.sConta & SEPARADOR & sCcl)
            
                'lErro = CF("ContaCcl_Saldo", iFilialEmpresa, sConta, sCcl, dtDataIni, iExerAnt, objPer, dSaldoIni)
                lErro = ContaCcl_Saldo_Sped(iFilialEmpresa, sConta, sCcl, dtDataIni, iExerAnt, objPer, dSaldoIni, alComando(25), alComando(26))
                If lErro <> SUCESSO Then gError 203194
                
                'lErro = CF("ContaCcl_Movimentacao", iFilialEmpresa, sConta, sCcl, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb)
                lErro = ContaCcl_Movimentacao_Sped(iFilialEmpresa, sConta, sCcl, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb, alComando(27))
                If lErro <> SUCESSO Then gError 203195
                
                If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
            
                    dSaldoCreAcum = Arredonda_Moeda(dSaldoCreAcum + dSaldoCre)
                    dSaldoDebAcum = Arredonda_Moeda(dSaldoDebAcum + dSaldoDeb)
                    dSaldoAcum = Arredonda_Moeda(dSaldoAcum + dSaldoIni)
                    dSaldoIniTotal = dSaldoIniTotal + dSaldoIni
                    
                    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                    lTotalRegBloco1 = lTotalRegBloco1 + 1
        
                    lErro = Mascara_MascararConta(sConta, sContaTela)
                    If lErro <> SUCESSO Then gError 203196
                    
                    lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 203197
                    
                    dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
                    dSaldoFimTotal = dSaldoFimTotal + dSaldoFim
                    
                    objSpedContabil.sRegistro = "|I155|" & sContaTela & "|" & sCclMascarado & "|" & Format(Abs(dSaldoIni), "fixed") & "|" & IIf(dSaldoIni > 0, "C", "D") & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & IIf(dSaldoFim > 0, "C", "D") & "|"
                    
                    lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, sCclMascarado, dSaldoIni, objSpedContabil)
                    If lErro <> SUCESSO Then gError 207328
                    
                    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
                    
'                    If Len(objSpedContabil.sRegistro) > 500 Then gError 203198
'
'                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'                    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
'                    If lErro <> SUCESSO Then gError 203199
            
                End If
            
            
                objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
                objTela.dValorAtual = objTela.dValorAtual + 1
            
                objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
            
                DoEvents
                
                If objTela.iCancelaBatch = CANCELA_BATCH Then
            
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
                
                    If vbMsgRes = vbYes Then gError 203268
                    
                    objTela.iCancelaBatch = 0
                        
                End If
            
                lErro = Comando_BuscarProximo(alComando(13))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203200
            
            
            Loop
        
                
            'lErro = CF("Conta_Saldo", iFilialEmpresa, sContaAtual, dtDataIni, iExerAnt, objPer, dSaldoIni)
            lErro = Conta_Saldo_Sped(iFilialEmpresa, sContaAtual, dtDataIni, iExerAnt, objPer, dSaldoIni, alComando(27), alComando(28))
            If lErro <> SUCESSO Then gError 203238
            
            'lErro = CF("Conta_Movimentacao", iFilialEmpresa, sContaAtual, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb)
            lErro = Conta_Movimentacao_Sped(iFilialEmpresa, sContaAtual, dtDataIni, dtDataFim, dSaldoCre, dSaldoDeb, alComando(29))
            If lErro <> SUCESSO Then gError 203239
        
            dSaldoIni = Arredonda_Moeda(dSaldoIni - dSaldoAcum)
            dSaldoCre = Arredonda_Moeda(dSaldoCre - dSaldoCreAcum)
            dSaldoDeb = Arredonda_Moeda(dSaldoDeb - dSaldoDebAcum)
            dSaldoFim = Arredonda_Moeda(dSaldoIni + dSaldoCre - dSaldoDeb)
            dSaldoIniTotal = dSaldoIniTotal + dSaldoIni
            dSaldoFimTotal = dSaldoFimTotal + dSaldoFim
        
           If dSaldoIni <> 0 Or dSaldoCre <> 0 Or dSaldoDeb <> 0 Then
        
                objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                lTotalRegBloco1 = lTotalRegBloco1 + 1
        
                lErro = Mascara_MascararConta(sContaAtual, sContaTela)
                If lErro <> SUCESSO Then gError 203246
        
                objSpedContabil.sRegistro = "|I155|" & sContaTela & "|" & "|" & Format(Abs(dSaldoIni), "fixed") & "|" & IIf(dSaldoIni > 0, "C", "D") & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & IIf(dSaldoFim > 0, "C", "D") & "|"
                
                sContaTelaUltimo = sContaTela
                dSaldoIniUltimo = dSaldoIni
                dSaldoFimUltimo = dSaldoFim
                lSeqUltimo = objSpedContabil.lSeq
                dSaldoCreUltimo = dSaldoCre
                dSaldoDebUltimo = dSaldoDeb
                
                objSpedRegistro.lCount = objSpedRegistro.lCount + 1
                
                lErro = SpedContabil1_Armazena_SldIni(colSpedContabil1, sContaTela, "", dSaldoIni, objSpedContabil)
                If lErro <> SUCESSO Then gError 207330
                
'                If Len(objSpedContabil.sRegistro) > 500 Then gError 203240
'
'                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'                objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
'                If lErro <> SUCESSO Then gError 203241
                
            End If
            
        
            lErro = Comando_BuscarProximo(alComando(22))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204747
        
        
            sContaAtual = tPlanoConta.sConta
            dSaldoAcum = 0
            dSaldoCreAcum = 0
            dSaldoDebAcum = 0
                
        
        Loop
        
'        dSaldoIni = dSaldoIniUltimo - dSaldoIniTotal
'        dSaldoFim = dSaldoFimUltimo - dSaldoIniTotal
'        dSaldoCre = dSaldoCreUltimo
'        dSaldoDeb = dSaldoDebUltimo
'
'        objSpedContabil.sRegistro = "|I155|" & sContaTelaUltimo & "|" & "|" & Format(Abs(dSaldoIni), "fixed") & "|" & IIf(dSaldoIni > 0, "C", "D") & "|" & Format(dSaldoDeb, "fixed") & "|" & Format(dSaldoCre, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & IIf(dSaldoFim > 0, "C", "D") & "|"
'
'        Set objSpedContabil1 = colSpedContabil1.Item(colSpedContabil1.Count)
        
'        objSpedContabil1.dSldIni = dSaldoIni
        
'        lErro = Comando_Executar(alComando(23), "UPDATE SpedContabil SET Registro = ? WHERE  NumIntSped = ? AND Seq = ?", _
'            objSpedContabil.sRegistro, objSpedContabil.lNumIntSped, lSeqUltimo)
'        If lErro <> SUCESSO Then gError 203241
        
        
        
    End If
    
    '************ registro I200 *******************
    
    Call objTela.SPED_Status("Processando Registro I200", "")
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(14), "SELECT  FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Data, Conta, Ccl, Historico, Valor FROM Lancamentos WHERE FilialEmpresa < 50 AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) ORDER BY FilialEmpresa, Origem,  Data, Doc, Seq ", _
        tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203201
    Else
        lErro = Comando_Executar(alComando(14), "SELECT  FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Data, Conta, Ccl, Historico, Valor FROM Lancamentos WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (Aglutinado = 0 OR Aglutinado = 1) ORDER BY Origem,  Data, Doc, Seq ", _
        tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203201
    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(14))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203202

    lDocAtual = -99999

    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = "I200"

    Set objSpedRegistroI250 = New ClassSpedRegistro

    objSpedRegistroI250.sTipo = "I250"

    colSpedRegistro.Add objSpedRegistroI250

    iFilialEmpresaAtual = tLanc.iFilialEmpresa
    sOrigemAtual = tLanc.sOrigem
    iExercicioAtual = tLanc.iExercicio
    iPeriodoLanAtual = tLanc.iPeriodoLan
    lDocAtual = tLanc.lDoc
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lSeqDoc = objSpedContabil.lSeq
    dtDataAtual = tLanc.dtData

    dValorDocPos = 0
    dValorDocNeg = 0
    
    Set colLanc = New Collection

    iX1 = 0
    Do While lErro = AD_SQL_SUCESSO
    
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
                
            
'            lErro = Comando_Executar(alComando(19), "SELECT  SUM(Valor) FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio= ? AND PeriodoLan = ? AND Doc = ? AND Valor > 0 AND (Aglutinado = 0 OR Aglutinado = 1) ", _
'            dValorDoc, iFilialEmpresaAtual, sOrigemAtual, iExercicioAtual, iPeriodoLanAtual, lDocAtual)
'            If lErro <> AD_SQL_SUCESSO Then gError 203203
'
'            lErro = Comando_BuscarPrimeiro(alComando(19))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203204
            
            sIdentificador = Format(iFilialEmpresaAtual, "00") & sOrigemAtual & Format(iExercicioAtual, "00") & Format(iPeriodoLanAtual, "00") & Format(lDocAtual, "000000000")
            
            objSpedContabil.sTipo = "I200"
            objSpedContabil.lSeq = objSpedContabil.lSeq + 1
            lTotalRegBloco1 = lTotalRegBloco1 + 1
            objSpedContabil.sRegistro = "|I200|" & sIdentificador & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & IIf(sOrigemAtual = "APE", "E", "N") & "|"
            
            
            objSpedRegistro.lCount = objSpedRegistro.lCount + 1
            
            If Len(objSpedContabil.sRegistro) > 500 Then gError 203205
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedContabil.lNumIntSped, lSeqDoc, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
            If lErro <> SUCESSO Then gError 203206
            
            
            For Each objLanc In colLanc
                
                objSpedContabil.sTipo = "I250"
                objSpedContabil.lSeq = objSpedContabil.lSeq + 1
                lTotalRegBloco1 = lTotalRegBloco1 + 1
            
                objSpedContabil.sRegistro = "|I250|" & objLanc.sConta & "|" & objLanc.sCcl & "|" & Format(Abs(objLanc.dValor), "fixed") & "|" & IIf(objLanc.dValor > 0, "C", "D") & "|||" & objLanc.sHistorico & "||"
            
                If Len(objSpedContabil.sRegistro) > 500 Then gError 203209
        
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
                If lErro <> SUCESSO Then gError 203210
            
                lErro = SpedContabil1_Atualiza_CredDeb(colSpedContabil1, objLanc.sConta, objLanc.sCcl, objLanc.dValor)
                If lErro <> SUCESSO Then gError 207331
            
            Next
            
            iFilialEmpresaAtual = tLanc.iFilialEmpresa
            sOrigemAtual = tLanc.sOrigem
            iExercicioAtual = tLanc.iExercicio
            iPeriodoLanAtual = tLanc.iPeriodoLan
            lDocAtual = tLanc.lDoc
            dtDataAtual = tLanc.dtData
            lSeqDoc = objSpedContabil.lSeq
            
            dValorDocPos = 0
            dValorDocNeg = 0
            
            Set colLanc = New Collection
            
        End If
    
        '************ registro I250 *******************
        
         Call objTela.SPED_Status("Processando Registro I250 - Lançamentos", "Doc: " & CStr(tLanc.lDoc) & " Data: " & Format(tLanc.dtData, "dd/mm/yyyy") & "Valor: " & Format(tLanc.dValor, "STANDARD"))
    
'        objSpedContabil.sTipo = "I250"
'        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
'        lTotalRegBloco1 = lTotalRegBloco1 + 1
        
        objSpedRegistroI250.lCount = objSpedRegistroI250.lCount + 1
        
        
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
        
'        objSpedContabil.sRegistro = "|I250|" & sContaTela & "|" & sCclMascarado & "|" & Format(Abs(Arredonda_Moeda(tLanc.dValor)), "fixed") & "|" & IIf(tLanc.dValor > 0, "C", "D") & "|||" & tLanc.sHistorico & "||"
        
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.sConta = sContaTela
        objLanc.sCcl = sCclMascarado
        objLanc.dValor = Arredonda_Moeda(tLanc.dValor)
        objLanc.sHistorico = tLanc.sHistorico
        
        colLanc.Add objLanc
        
        
        
'        If Len(objSpedContabil.sRegistro) > 500 Then gError 203209
'
'        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
'        If lErro <> SUCESSO Then gError 203210
    
        objTela.TotReg.Caption = CStr(CLng(objTela.TotReg.Caption) + 1)
        objTela.dValorAtual = objTela.dValorAtual + 1
    
        objTela.ProgressBar1.Value = CInt((objTela.dValorAtual / objTela.dValorTotal) * 100)
    
        DoEvents
        
        If objTela.iCancelaBatch = CANCELA_BATCH Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_DIARIO")
        
            If vbMsgRes = vbYes Then gError 203264
            
            objTela.iCancelaBatch = 0
                
        End If
    
    
        lErro = Comando_BuscarProximo(alComando(14))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203211
    
    Loop
    
'    If dValorDoc <> 0 Then
    
'        sIdentificador = Format(iFilialEmpresaAtual, "00") & sOrigemAtual & Format(iExercicioAtual, "00") & Format(iPeriodoLanAtual, "00") & Format(lDocAtual, "000000000")
'
'        objSpedContabil.sTipo = "I200"
'        lTotalRegBloco1 = lTotalRegBloco1 + 1
'        objSpedContabil.sRegistro = "|I200|" & sIdentificador & "|" & Format(tLanc.dtData, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & IIf(sOrigemAtual = "APE", "E", "N") & "|"
'
'
'        objSpedRegistro.lCount = objSpedRegistro.lCount + 1
'
'        If Len(objSpedContabil.sRegistro) > 500 Then gError 203205
'
'        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'        objSpedContabil.lNumIntSped, lSeqDoc, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
'        If lErro <> SUCESSO Then gError 203206
'   End If
    
    
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
        
        
        
    
'            lErro = Comando_Executar(alComando(19), "SELECT  SUM(Valor) FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio= ? AND PeriodoLan = ? AND Doc = ? AND Valor > 0 AND (Aglutinado = 0 OR Aglutinado = 1) ", _
'            dValorDoc, iFilialEmpresaAtual, sOrigemAtual, iExercicioAtual, iPeriodoLanAtual, lDocAtual)
'            If lErro <> AD_SQL_SUCESSO Then gError 203203
'
'            lErro = Comando_BuscarPrimeiro(alComando(19))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203204
    
    sIdentificador = Format(iFilialEmpresaAtual, "00") & sOrigemAtual & Format(iExercicioAtual, "00") & Format(iPeriodoLanAtual, "00") & Format(lDocAtual, "000000000")
    
    objSpedContabil.sTipo = "I200"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    objSpedContabil.sRegistro = "|I200|" & sIdentificador & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(Arredonda_Moeda(dValorDoc), "fixed") & "|" & IIf(sOrigemAtual = "APE", "E", "N") & "|"
    
    
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203205
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, lSeqDoc, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203206
    
    
    For Each objLanc In colLanc
        
        objSpedContabil.sTipo = "I250"
        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
        lTotalRegBloco1 = lTotalRegBloco1 + 1
    
        objSpedContabil.sRegistro = "|I250|" & objLanc.sConta & "|" & objLanc.sCcl & "|" & Format(Abs(objLanc.dValor), "fixed") & "|" & IIf(objLanc.dValor > 0, "C", "D") & "|||" & objLanc.sHistorico & "||"
    
        If Len(objSpedContabil.sRegistro) > 500 Then gError 203209

        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
        If lErro <> SUCESSO Then gError 203210
    
        lErro = SpedContabil1_Atualiza_CredDeb(colSpedContabil1, objLanc.sConta, objLanc.sCcl, objLanc.dValor)
        If lErro <> SUCESSO Then gError 207331
    
    Next
    
    For Each objSpedContabil1 In colSpedContabil1
    
        Set objSpedContabil = objSpedContabil1.objSpedContabil
        
        dSaldoFim = objSpedContabil1.dSldIni + objSpedContabil1.dCred - objSpedContabil1.dDeb
        
        objSpedContabil.sRegistro = "|I155|" & objSpedContabil1.sConta & "|" & objSpedContabil1.sCcl & "|" & Format(Abs(objSpedContabil1.dSldIni), "fixed") & "|" & IIf(objSpedContabil1.dSldIni > 0, "C", "D") & "|" & Format(objSpedContabil1.dDeb, "fixed") & "|" & Format(objSpedContabil1.dCred, "fixed") & "|" & Format(Abs(dSaldoFim), "fixed") & "|" & IIf(dSaldoFim > 0, "C", "D") & "|"
                
        If Len(objSpedContabil.sRegistro) > 500 Then gError 203240

        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
        If lErro <> SUCESSO Then gError 203241
    
    
    Next
    
    '************ registro I990 *******************

    Call objTela.SPED_Status("Processando Registro I990", "")

    objSpedContabil.sTipo = "I990"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    lTotalRegBloco1 = lTotalRegBloco1 + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    objSpedContabil.sRegistro = "|I990|" & lTotalRegBloco1 & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203212
    
    '************ registro J001 *******************
  
    Call objTela.SPED_Status("Processando Registro J001", "")
  
    objSpedContabil.sTipo = "J001"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    objSpedContabil.sRegistro = "|J001|0|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203213
    
    '************ registro J900 *******************

    Call objTela.SPED_Status("Processando Registro J900", "")

    objSpedContabil.sTipo = "J900"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    objSpedContabil.sRegistro = "|J900|TERMO DE ENCERRAMENTO|" & lNumOrd & "|Escrituração Contábil Digital do Livro Diário Geral|"
    objSpedContabil.sRegistro = objSpedContabil.sRegistro & objFilialEmpresa.sNome & "||" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203214
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203215
    
    '************ registro J930 *******************

    Call objTela.SPED_Status("Processando Registro J930", "")

    objSpedContabil.sTipo = "J930"
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    sDescQualiSigCTB = String(STRING_DNRC_DESC, 0)
    
    lErro = Comando_Executar(alComando(15), "SELECT  Descricao FROM DNRC WHERE Codigo = ?", _
    sDescQualiSigCTB, objFilialEmpresa.sCodQualiSigCTB)
    If lErro <> AD_SQL_SUCESSO Then gError 203216
    
    lErro = Comando_BuscarPrimeiro(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203217
    
    objSpedContabil.sRegistro = "|J930|" & objFilialEmpresa.sSignatarioCTB & "|" & objFilialEmpresa.sCPFSignatarioCTB & "|" & sDescQualiSigCTB & "|" & objFilialEmpresa.sCodQualiSigCTB & "||"
    
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203218
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203219
    
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    objSpedContabil.sRegistro = "|J930|" & objFilialEmpresa.sContador & "|" & objFilialEmpresa.sCPFContador & "|" & "Contador" & "|" & "900" & "|" & objFilialEmpresa.sCRCContador & "|"
    
    If Len(objSpedContabil.sRegistro) > 500 Then gError 203248
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203249
    
    
    '************ registro J990 *******************
    
    Call objTela.SPED_Status("Processando Registro J990", "")
    
    objSpedContabil.sTipo = "J990"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    objSpedContabil.sRegistro = "|J990|5|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203220

    '************ registro 9001 *******************
    
    Call objTela.SPED_Status("Processando Registro 9001", "")
   
    objSpedContabil.sTipo = "9001"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = objSpedContabil.sTipo
    objSpedRegistro.lCount = objSpedRegistro.lCount + 1
    
    objSpedContabil.sRegistro = "|9001|0|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203221

    '************ registro 9900 *******************
    
    Call objTela.SPED_Status("Processando Registro 9900", "")
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = "9900"
    objSpedRegistro.lCount = colSpedRegistro.Count + 2
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = "9990"
    objSpedRegistro.lCount = 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    colSpedRegistro.Add objSpedRegistro
    
    objSpedRegistro.sTipo = "9999"
    objSpedRegistro.lCount = 1
    
    
    objSpedContabil.sTipo = "9900"
    
    iX1 = 0
    For Each objSpedRegistro In colSpedRegistro
    
        iX1 = iX1 + 1
    
        Call objTela.SPED_Status("Processando Registro 9900", CStr(iX1) & SEPARADOR & objSpedContabil.sTipo)
    
        objSpedContabil.lSeq = objSpedContabil.lSeq + 1
        
        objSpedContabil.sRegistro = "|9900|" & objSpedRegistro.sTipo & "|" & objSpedRegistro.lCount & "|"
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
        If lErro <> SUCESSO Then gError 203222

    Next
    
    '************ registro 9990 *******************
    
    Call objTela.SPED_Status("Processando Registro 9990", "")
    
    objSpedContabil.sTipo = "9990"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    objSpedContabil.sRegistro = "|9990|" & colSpedRegistro.Count + 3 & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203223
    
    '************ registro 9999 *******************
    
    Call objTela.SPED_Status("Processando Registro 9999", "")
    
    objSpedContabil.sTipo = "9999"
    objSpedContabil.lSeq = objSpedContabil.lSeq + 1
    
    objSpedContabil.sRegistro = "|9999|" & objSpedContabil.lSeq & "|"
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedContabil( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedContabil.lNumIntSped, objSpedContabil.lSeq, objSpedContabil.iFilialEmpresa, objSpedContabil.sTipo, objSpedContabil.sRegistro)
    If lErro <> SUCESSO Then gError 203224
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'I030'", _
    sRegistro, objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203225
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203226

    iPos = 1

    For iIndice = 1 To 4
        iPos = iPos + 1
        iPos = InStr(iPos, sRegistro, "|")
    Next
    
    objSpedContabil.sRegistro = Mid(sRegistro, 1, iPos) & objSpedContabil.lSeq & Mid(sRegistro, iPos + 1)
    
    lErro = Comando_Executar(alComando(17), "UPDATE SpedContabil SET Registro = ? WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'I030'", _
    objSpedContabil.sRegistro, objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 203227
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'J900'", _
    sRegistro, objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203228
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203229

    iPos = 1

    For iIndice = 1 To 5
        iPos = iPos + 1
        iPos = InStr(iPos, sRegistro, "|")
    Next
    
    objSpedContabil.sRegistro = Mid(sRegistro, 1, iPos) & objSpedContabil.lSeq & Mid(sRegistro, iPos + 1)
    
    lErro = Comando_Executar(alComando(17), "UPDATE SpedContabil SET Registro = ? WHERE NumIntSped = ? AND FilialEmpresa = ? And Tipo = 'J900'", _
    objSpedContabil.sRegistro, objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 203230

    Call objTela.SPED_Status("Gerando arquivo", "")

    Open sDiretorio For Output As #1
    
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(16), "SELECT  Registro FROM SpedContabil WHERE NumIntSped = ? AND FilialEmpresa = ? ORDER BY Seq", _
    sRegistro, objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
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

    lErro = Comando_Executar(alComando(18), "DELETE FROM SpedContabil WHERE  NumIntSped = ? AND FilialEmpresa = ?", _
    objSpedContabil.lNumIntSped, objSpedContabil.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 203234

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 203235
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_DIARIO_SPED_SUCESSO", sDiretorio)
    
    Gera_Sped_Contabil_Diario = SUCESSO
    
    Exit Function
    
Erro_Gera_Sped_Contabil_Diario:

    Gera_Sped_Contabil_Diario = gErr
    Resume Next
    Select Case gErr
    
        Case 203059, 203060, 203068, 203069, 203115, 203119, 203121, 203143, 203144, 203151, 203152, 203159, 203168, 203179, 203186 To 203188, 203194 To 203197, 203207, 203208, 203238, 203239, 203242, 203243, 203246, 203247, 203264 To 203268, 207328, 207329, 207330
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, objCidades.sDescricao)

        Case 203123, 203128, 203131, 203133, 203135, 203137, 203145, 203160, 203180, 203189, 203198, 203205, 203209, 203214, 203218, 203240, 203244, 203248
            Call Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_MAIOR_TAM_MAXIMO", gErr)

        Case 203124, 203125, 203129, 203132, 203134, 203136, 203138, 203146, 203153, 203154, 203157, 203161, 203162, 203165, 203169, 203181, 203183, 203190, 203199, 203206, 203210, 203212, 203213, 203215, 203219 To 203224, 203241, 203245, 203249, 204748
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDCONTABIL", gErr)

        Case 203126, 203127, 203130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDCADASTRO", gErr)

        Case 203139, 203140, 203172, 203184, 203185, 203191, 203254, 203255, 203260, 203261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)

        Case 203141, 203142
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)

        Case 203147, 203148
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 203149, 203150, 203158, 203166, 203167, 203170, 203175, 203176, 203182, 203200, 203256, 203257
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", gErr)

        Case 203155, 203156, 203163, 203164, 203171
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
    
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203236)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Call Transacao_Rollback
   
    Exit Function
    
End Function

Function Conta_Saldo(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 4) As Long, iIndice As Integer
Dim iExercicio As Integer
Dim dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer
Dim dValMvDia As Double

On Error GoTo Erro_Conta_Saldo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203058
    Next

    'percorrer registros gerando os saldos na data para as diversas contas
    dSaldo = 0
    iMes = Month(dtData)
    
    lErro = Comando_Executar(alComando(2), "SELECT Exercicio, SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta = ? ORDER BY Exercicio", _
        iExercicio, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 203061
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
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
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203063
    
    Loop
    
    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Cre-Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Data >= ? AND Data < ? AND Conta = ?", _
            dValMvDia, iFilialEmpresa, objPer.dtDataInicio, dtData, sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 203064
                
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203065
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Arredonda_Moeda(dSaldo + dValMvDia, 2)
        
    End If
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Conta_Saldo = SUCESSO
     
    Exit Function
    
Erro_Conta_Saldo:

    Conta_Saldo = gErr
     
    Select Case gErr
          
        Case 203058
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 203061 To 203063
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case 203064 To 203065
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203066)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function


Function ContaCcl_Saldo(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 4) As Long, iIndice As Integer
Dim iExercicio As Integer
Dim dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer
Dim dValMvDia As Double

On Error GoTo Erro_ContaCcl_Saldo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203067
    Next

    'percorrer registros gerando os saldos na data para as diversas contas
    dSaldo = 0
    iMes = Month(dtData)
    
    lErro = Comando_Executar(alComando(2), "SELECT Exercicio, SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta = ? AND Ccl = ? ORDER BY Exercicio", _
        iExercicio, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 203070
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
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
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203072
    
    Loop
    
    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Cre-Deb) FROM MvDiaCcl WHERE FilialEmpresa = ? AND Data >= ? AND Data < ? AND Conta = ? AND Ccl = ?", _
            dValMvDia, iFilialEmpresa, objPer.dtDataInicio, dtData, sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 203073
                
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203074
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Arredonda_Moeda(dSaldo + dValMvDia, 2)
        
    End If
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    ContaCcl_Saldo = SUCESSO
     
    Exit Function
    
Erro_ContaCcl_Saldo:

    ContaCcl_Saldo = gErr
     
    Select Case gErr
          
        Case 203067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 203070 To 203072
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
        
        Case 203073 To 203074
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203075)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Conta_Movimentacao(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 1) As Long, iIndice As Integer

On Error GoTo Erro_Conta_Movimentacao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203076
    Next

    
    dSaldoCre = 0
    dSaldoDeb = 0
    
    'obtem valor movimentado no periodo (mes) até a data
    lErro = Comando_Executar(alComando(1), "SELECT SUM(Cre), SUM(Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND Conta = ?", _
        dSaldoCre, dSaldoDeb, iFilialEmpresa, dtDataIni, dtDataFim, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 203077
            
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203078
        
        
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Conta_Movimentacao = SUCESSO
     
    Exit Function
    
Erro_Conta_Movimentacao:

    Conta_Movimentacao = gErr
     
    Select Case gErr
          
        Case 203076
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 203077, 203078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203079)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function


Function ContaCcl_Movimentacao(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 1) As Long, iIndice As Integer

On Error GoTo Erro_ContaCcl_Movimentacao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203080
    Next

    dSaldoCre = 0
    dSaldoDeb = 0

    'obtem valor movimentado no periodo (mes) até a data
    lErro = Comando_Executar(alComando(1), "SELECT SUM(Cre) , SUM(Deb) FROM MvDiaCcl WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND Conta = ? AND Ccl = ?", _
        dSaldoCre, dSaldoDeb, iFilialEmpresa, dtDataIni, dtDataFim, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 203081
            
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203082
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    ContaCcl_Movimentacao = SUCESSO
     
    Exit Function
    
Erro_ContaCcl_Movimentacao:

    ContaCcl_Movimentacao = gErr
     
    Select Case gErr
          
        Case 203080
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 203081, 203082
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203083)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Periodo_Le_Sped(ByVal dtData As Date, ByVal objPeriodo As ClassPeriodo, ByVal lComando As Long) As Long
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

Function ContaCcl_Saldo_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double, lComando1 As Long, lComando2 As Long) As Long
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

Function ContaCcl_Movimentacao_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double, lComando As Long) As Long
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

Function Conta_Saldo_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtData As Date, ByVal iExerAnt As Integer, ByVal objPer As ClassPeriodo, dSaldo As Double, lComando1 As Long, lComando2 As Long) As Long
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

Function Conta_Movimentacao_Sped(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dSaldoCre As Double, dSaldoDeb As Double, lComando As Long) As Long
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

Function SpedContabil1_Armazena_SldIni(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dSldIni As Double, ByVal objSpedContabil As ClassSpedContabil) As Long

    Dim objSpedContabil1 As New ClassSpedContabil1
    Dim objSpedContabilAux As New ClassSpedContabil
    
    objSpedContabilAux.iFilialEmpresa = objSpedContabil.iFilialEmpresa
    objSpedContabilAux.lNumIntSped = objSpedContabil.lNumIntSped
    objSpedContabilAux.lSeq = objSpedContabil.lSeq
    objSpedContabilAux.lSeq = objSpedContabil.lSeq
    objSpedContabilAux.sTipo = objSpedContabil.sTipo
    
    Set objSpedContabil1.objSpedContabil = objSpedContabil
    objSpedContabil1.sConta = sContaTela
    objSpedContabil1.sCcl = sCclMascarado
    objSpedContabil1.dSldIni = dSldIni
    
    colSpedContabil1.Add objSpedContabil1
    
    SpedContabil1_Armazena_SldIni = SUCESSO

End Function

Function SpedContabil1_Atualiza_CredDeb(ByVal colSpedContabil1 As Collection, ByVal sContaTela As String, ByVal sCclMascarado As String, ByVal dValor As Double) As Long

    Dim objSpedContabil1 As New ClassSpedContabil1
    
    For Each objSpedContabil1 In colSpedContabil1
    
        If objSpedContabil1.sConta = sContaTela And objSpedContabil1.sCcl = sCclMascarado Then
            If dValor > 0 Then
                objSpedContabil1.dCred = objSpedContabil1.dCred + dValor
            Else
                objSpedContabil1.dDeb = objSpedContabil1.dDeb - dValor
            End If
            
            Exit For
        
        End If
        
    Next
    
    SpedContabil1_Atualiza_CredDeb = SUCESSO

End Function

