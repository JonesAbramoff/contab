VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISSpedECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Gera_Sped_ECF(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long
Dim bArqAberto As Boolean

On Error GoTo Erro_Gera_Sped_ECF

    bArqAberto = False

    'Obtem totalizadores, grava SpedDocFiscais e lê dados como filial empresa e outros que serão usados no processamento
    Call objSped.objTela.SPED_Status("Obtendo dados inicias", "")
    lErro = Sped_Obtem_DadosInicias(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objSped.lFreeFile = FreeFile()

    Open objSped.sDiretorio For Output As #objSped.lFreeFile
    bArqAberto = True

    'BLOCO 0: Abertura, Identificação e Referências
    Call objSped.objTela.SPED_Status("Processando Bloco 0", "")
    lErro = Processa_Bloco_0(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'BLOCO J: Plano de Contas e Mapeamento
    Call objSped.objTela.SPED_Status("Processando Bloco J", "")
    lErro = Processa_Bloco_J(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'BLOCO 9: Controle e Encerramento do Arquivo Digital
    Call objSped.objTela.SPED_Status("Processando Bloco 9", "")
    lErro = Processa_Bloco_9(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Close #objSped.lFreeFile
    bArqAberto = False

    Call Rotina_Aviso(vbOKOnly, "AVISO_SPED_ECF_SUCESSO", objSped.sDiretorio)

    Gera_Sped_ECF = SUCESSO

    Exit Function

Erro_Gera_Sped_ECF:

    Gera_Sped_ECF = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213890)

    End Select
    
    If bArqAberto Then Close #objSped.lFreeFile

    Exit Function

End Function

Private Function Sped_Obtem_DadosInicias(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long
Dim iCodVersao As Integer
Dim lTransacao As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais, lContador As Long
Dim lModelo As Long, iTipoModelo As Integer

On Error GoTo Erro_Sped_Obtem_DadosInicias

    'Abertura de Comando
    For iIndice = 0 To objSped.NUM_MAX_COMANDOS
        objSped.alComando(iIndice) = Comando_Abrir()
        If objSped.alComando(iIndice) = 0 Then gError 213891
    Next

    'Obtem o total de registros que vão atualizar a barra de progressão
    lErro = Registros_Totaliza(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call objSped.objTela.Inicia_Processo(objSped.lNumTotalPrevisto)

    'Obtém a versão do layout
    lErro = Comando_Executar(objSped.ProxComando, "SELECT Codigo From SpedECFVersaoLeiaute WHERE DataInicio <= ? ORDER BY DataInicio DESC ", iCodVersao, objSped.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 213892

    lErro = Comando_BuscarPrimeiro(objSped.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213893

    'versao nao cadastrada
    If lErro <> AD_SQL_SUCESSO Then gError 213894

    objFilialEmpresa.iCodFilial = giFilialEmpresa

    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM

    objSped.iCodVersao = iCodVersao

    objSped.objSpedFiscal.iFilialEmpresa = objSped.iFilialEmpresa
    objSped.objSpedFiscal.lNumIntSped = objSped.lNumIntSped
    
    Set objSped.objFilialEmpresa = objFilialEmpresa
    
    ' descobre o plano de contas referencial que sera usado
    lErro = Comando_Executar(objSped.ProxComando, "SELECT Codigo, Tipo FROM PlanoContaRefModelo WHERE Oficial = 1 AND AnoVigencia <= ? ORDER BY AnoVigencia DESC", _
    lModelo, iTipoModelo, Year(objSped.dtDataFim))
    If lErro <> AD_SQL_SUCESSO Then gError 213895

    lErro = Comando_BuscarPrimeiro(objSped.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213896
    
    If lErro <> AD_SQL_SUCESSO Then gError 213897
    
    objSped.lModelo = lModelo
    objSped.iTipoModelo = iTipoModelo

    Sped_Obtem_DadosInicias = SUCESSO

    Exit Function

Erro_Sped_Obtem_DadosInicias:

   Sped_Obtem_DadosInicias = gErr

   Select Case gErr

        Case 213891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213892, 213893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDECFVERSAOLEIAUTE", gErr)

        Case 213894
            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_SEM_VERSAO", gErr, Format(objSped.dtDataFim, "dd/mm/yyyy"))

        Case 213895, 213896
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 213897
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTAREFMODELO_NAO_CADASTRADO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213898)

    End Select

End Function

Private Function Registros_Totaliza(ByVal objSped As ClassSpedECFSel) As Long
'Verifica os registros com grande processamento para ajustar a barra de progressão

Dim lErro As Long
Dim lContador As Long

On Error GoTo Erro_Registros_Totaliza

    Call objSped.objTela.SPED_Status("Totalizando Registros", "")
    
    lErro = Comando_Executar(objSped.ProxComando, "SELECT COUNT(*) FROM PlanoConta ", lContador)
    If lErro <> AD_SQL_SUCESSO Then gError 213899

    lErro = Comando_BuscarPrimeiro(objSped.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213900
    
    'É uma previsão mais ou menos que vai ser ajustada no decorrer do processo
    objSped.lNumTotalPrevisto = lContador * 2

    Registros_Totaliza = SUCESSO

    Exit Function

Erro_Registros_Totaliza:

   Registros_Totaliza = gErr

   Select Case gErr

        Case 213899, 213900
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213901)

    End Select

End Function

Private Sub Registro_Adiciona_Count(ByVal sTipo As String, ByVal colSpedRegistro As Collection, ByVal iCount As Integer)
Dim objSpedRegistro As New ClassSpedRegistro
Dim bAchou As Boolean
    If iCount > 0 Then
        bAchou = False
        For Each objSpedRegistro In colSpedRegistro
            If objSpedRegistro.sTipo = sTipo Then
                objSpedRegistro.lCount = objSpedRegistro.lCount + iCount
                bAchou = True
                Exit For
            End If
        Next
        If bAchou = False Then
            Set objSpedRegistro = New ClassSpedRegistro
            objSpedRegistro.sTipo = sTipo
            objSpedRegistro.lCount = iCount
            colSpedRegistro.Add objSpedRegistro
        End If
    End If
End Sub

Private Function Registro_Insere(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long
Dim objSpedFiscal As ClassSpedFiscal
Dim sAux As String, iSeq As Long, sReg As String, sErro As String

On Error GoTo Erro_Registro_Insere

    Set objSpedFiscal = objSped.objSpedFiscal

    If left(objSpedFiscal.sTipo, 2) <> "99" Then Call Registro_Adiciona_Count(objSped.objSpedFiscal.sTipo, objSped.colSpedRegistro, 1)

    Print #objSped.lFreeFile, objSpedFiscal.sRegistro
    
    lErro = objSped.objTela.Processa_Item(True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objSped.lNumTotalProcessados = objSped.lNumTotalProcessados + 1 'Total de processamentos (3 por registro)
    objSped.lNumTotalReal = objSped.lNumTotalReal + 1 'Total real de registros processados

    Registro_Insere = SUCESSO

    Exit Function

Erro_Registro_Insere:

   Registro_Insere = gErr

   Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213902)

    End Select

End Function

Private Function Processa_Bloco_0(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Bloco_0

    '========================================================================================
    '=============================== INÍCIO DO BLOCO 0 ======================================
    '================== BLOCO 0: ABERTURA, IDENTIFICAÇÃO E REFERÊNCIAS. =====================
    '========================================================================================

    'REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA
    Call objSped.objTela.SPED_Status("Processando Registro 0000", "")
    lErro = Processa_Registro_0000(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0001: ABERTURA DO BLOCO 0
    Call objSped.objTela.SPED_Status("Processando Registro 0001", "")
    lErro = Processa_Registro_0001(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'REGISTRO 0010: PARÂMETROS DE TRIBUTAÇÃO
    Call objSped.objTela.SPED_Status("Processando Registro 0010", "")
    lErro = Processa_Registro_0010(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0020: PARÂMETROS COMPLEMENTARES
    Call objSped.objTela.SPED_Status("Processando Registro 0020", "")
    lErro = Processa_Registro_0020(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0030: DADOS CADASTRAIS
    Call objSped.objTela.SPED_Status("Processando Registro 0030", "")
    lErro = Processa_Registro_0030(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
    Call objSped.objTela.SPED_Status("Processando Registro 0990", "")
    lErro = Processa_Registro_0990(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Processa_Bloco_0 = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_0:

    Processa_Bloco_0 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213903)
     
    End Select
    
End Function

Private Function Processa_Bloco_J(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Bloco_J
    
    'REGISTRO J001: ABERTURA DO BLOCO J
    Call objSped.objTela.SPED_Status("Processando Registro J001", "")
    lErro = Processa_Registro_J001(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO J050: PLANO DE CONTAS DO CONTRIBUINTE
    'REGISTRO J051: PLANO DE CONTAS REFERENCIAL
    Call objSped.objTela.SPED_Status("Processando Registro J050", "")
    lErro = Processa_Registro_J050(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO J100: CENTRO DE CUSTOS
    Call objSped.objTela.SPED_Status("Processando Registro J100", "")
    lErro = Processa_Registro_J100(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO J990: ENCERRAMENTO DO BLOCO J
    Call objSped.objTela.SPED_Status("Processando Registro J990", "")
    lErro = Processa_Registro_J990(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Processa_Bloco_J = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_J:

    Processa_Bloco_J = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213904)
     
    End Select
    
End Function

Private Function Processa_Bloco_9(ByVal objSped As ClassSpedECFSel) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Bloco_9

    '========================================================================================
    '=============================== INÍCIO DO BLOCO 9 ======================================
    '================ BLOCO 9: CONTROLE E ENCERRAMENTO DO ARQUIVO DIGITAL ===================
    '========================================================================================
    
    'REGISTRO 9001: ABERTURA DO BLOCO 9
    Call objSped.objTela.SPED_Status("Processando Registro 9001", "")
    lErro = Processa_Registro_9001(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9900: REGISTROS DO ARQUIVO
    Call objSped.objTela.SPED_Status("Processando Registro 9900", "")
    lErro = Processa_Registro_9900(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
    Call objSped.objTela.SPED_Status("Processando Registro 9990", "")
    lErro = Processa_Registro_9990(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
    Call objSped.objTela.SPED_Status("Processando Registro 9999", "")
    lErro = Processa_Registro_9999(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Processa_Bloco_9 = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_9:

    Processa_Bloco_9 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213905)
     
    End Select
    
End Function

Private Function Processa_Registro_0000(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA
Dim lErro As Long
Dim objCidades As New ClassCidades
Dim objSpedRegistro As ClassSpedRegistro
Dim sRazaoSocial As String
Dim lComandoDic As Long
Dim sIE As String, sIND_SIT_ESP As String, sIND_ATIV As String
Dim lComando As Long, sRazaoSocialAux1 As String, sRazaoSocialAux2 As String

On Error GoTo Erro_Processa_Registro_0000

    '******** registro 0000 *****************
    objSped.objSpedFiscal.sTipo = "0000"
    objSped.objSpedFiscal.lSeqPai = 0
    
    'abre conexao com o BD
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 213906

    lComando = objSped.ProxComando
    
    sRazaoSocialAux1 = String(STRING_NOME_EMPRESA, 0)
    
    'le a empresa
    lErro = Comando_Executar(lComandoDic, "SELECT Nome FROM Empresas WHERE Codigo = ? ", sRazaoSocialAux1, glEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213907
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213908
          
    sRazaoSocial = sRazaoSocialAux1
    
    sRazaoSocialAux2 = String(255, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM AdmConfig WHERE Codigo = ? AND FilialEmpresa = ? ", sRazaoSocialAux2, "EMPRESA_RAZAO_SOCIAL", objSped.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213909
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213910
    
    If lErro = AD_SQL_SUCESSO Then sRazaoSocial = sRazaoSocialAux2
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0000")
    Call objSped.Reg_Campo_add(2, "NOME_ESC", "LECF")
    Call objSped.Reg_Campo_add(3, "COD_VER", Format(objSped.iCodVersao, "0000"))
    Call objSped.Reg_Campo_add(4, "CNPJ", IIf(Len(objSped.objFilialEmpresa.sCgc) = 14, objSped.objFilialEmpresa.sCgc, ""))
    Call objSped.Reg_Campo_add(5, "NOME", SFF_Texto(sRazaoSocial, 100))
    Call objSped.Reg_Campo_add(6, "IND_SIT_INI_PER", "0")
    Call objSped.Reg_Campo_add(7, "SIT_ESPECIAL", "0")
    Call objSped.Reg_Campo_add(8, "PAT_REMAN_CIS", "")
    Call objSped.Reg_Campo_add(9, "DT_SIT_ESP", "")
    Call objSped.Reg_Campo_add(10, "DT_INI", SFF_Data(objSped.dtDataIni))
    Call objSped.Reg_Campo_add(11, "DT_FIN", SFF_Data(objSped.dtDataFim))
    Call objSped.Reg_Campo_add(12, "RETIFICADORA", "N")
    Call objSped.Reg_Campo_add(13, "NUM_REC", "")
    Call objSped.Reg_Campo_add(14, "TIP_ECF", "0")
    Call objSped.Reg_Campo_add(15, "COD_SCP", "")
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Comando_Fechar(lComandoDic)

    Processa_Registro_0000 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0000:

   Processa_Registro_0000 = gErr

   Select Case gErr
    
        Case 213906
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213907 To 213910
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESAS", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213911)
     
    End Select

    Call Comando_Fechar(lComandoDic)

    Exit Function
    
End Function

Private Function Processa_Registro_0001(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0001: ABERTURA DO BLOCO 0
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_0001

    '******** registro 0001 *****************
    objSped.objSpedFiscal.sTipo = "0001"
    objSped.objSpedFiscal.lSeqPai = 0
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0001")
    Call objSped.Reg_Campo_add(2, "IND_MOV", "0")
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0001:

   Processa_Registro_0001 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213912)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0010(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0010: PARÂMETROS DE TRIBUTAÇÃO
Dim lErro As Long
Dim sFormaTrib As String, sFormaTribPer As String

On Error GoTo Erro_Processa_Registro_0010

    '******** registro 0010 *****************
    objSped.objSpedFiscal.sTipo = "0010"
    objSped.objSpedFiscal.lSeqPai = 0
    
'    Forma de Tributação:
'    1 – Lucro Real.
'    2 – Lucro Real/Arbitrado.
'    3 – Lucro Presumido/Real.
'    4 – Lucro Presumido/Real/Arbitrado.
'    5 – Lucro Presumido.
'    6 – Lucro Arbitrado.
'    7 – Lucro Presumido/Arbitrado.
'    8 – Imune do IRPJ.
'    9 – Isenta do IRPJ.
    If objSped.objFilialEmpresa.iLucroPresumido = MARCADO Then
        sFormaTrib = "5"
        sFormaTribPer = "PPPP"
    Else
        sFormaTrib = "1"
        sFormaTribPer = "RRRR"
    End If
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0010")
    Call objSped.Reg_Campo_add(2, "HASH_ECF_ANTERIOR", "")
    Call objSped.Reg_Campo_add(3, "OPT_REFIS", "N")
    Call objSped.Reg_Campo_add(4, "OPT_PAES", "N")
    Call objSped.Reg_Campo_add(5, "FORMA_TRIB", sFormaTrib)
    Call objSped.Reg_Campo_add(6, "FORMA_APUR", "A")
    Call objSped.Reg_Campo_add(7, "COD_QUALIF_PJ", "01")
    Call objSped.Reg_Campo_add(8, "FORMA_TRIB_PER", sFormaTribPer)
    Call objSped.Reg_Campo_add(9, "MES_BAL_RED", "EEEEEEEEEEEE")
    Call objSped.Reg_Campo_add(10, "TIP_ESC_PRE", "")
    Call objSped.Reg_Campo_add(11, "TIP_ENT", "")
    Call objSped.Reg_Campo_add(12, "FORMA_APUR_I", "")
    Call objSped.Reg_Campo_add(13, "APUR_CSLL", "")
    Call objSped.Reg_Campo_add(14, "OPT_EXT_RTT", "N")
    Call objSped.Reg_Campo_add(15, "DIF_FCONT", "")
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0010:

   Processa_Registro_0010 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213913)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0020(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0020: PARÂMETROS COMPLEMENTARES
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_0020

    '******** registro 0020 *****************
    objSped.objSpedFiscal.sTipo = "0020"
    objSped.objSpedFiscal.lSeqPai = 0
       
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0020")
    Call objSped.Reg_Campo_add(2, "IND_ALIQ_CSLL", "N")
    Call objSped.Reg_Campo_add(3, "IND_QTE_SCP", "0")
    Call objSped.Reg_Campo_add(4, "IND_ADM_FUN_CLU", "N")
    Call objSped.Reg_Campo_add(5, "IND_PART_CONS", "N")
    Call objSped.Reg_Campo_add(6, "IND_OP_EXT", "N")
    Call objSped.Reg_Campo_add(7, "IND_OP_VINC", "N")
    Call objSped.Reg_Campo_add(8, "IND_PJ_ENQUAD", "N")
    Call objSped.Reg_Campo_add(9, "IND_PART_EXT", "N")
    Call objSped.Reg_Campo_add(10, "IND_ATIV_RURAL", "N")
    Call objSped.Reg_Campo_add(11, "IND_LUC_EXP", "N")
    Call objSped.Reg_Campo_add(12, "IND_RED_ISEN", "N")
    Call objSped.Reg_Campo_add(13, "IND_FIN", "N")
    Call objSped.Reg_Campo_add(14, "IND_DOA_ELEIT", "N")
    Call objSped.Reg_Campo_add(15, "IND_PART_COLIG", "N")
    Call objSped.Reg_Campo_add(16, "IND_VEND_EXP", "N")
    Call objSped.Reg_Campo_add(17, "IND_ REC_EXT", "N")
    Call objSped.Reg_Campo_add(18, "IND_ATIV_EXT", "N")
    Call objSped.Reg_Campo_add(19, "IND_COM_EXP", "N")
    Call objSped.Reg_Campo_add(20, "IND_PGTO_EXT", "N")
    Call objSped.Reg_Campo_add(21, "IND_E-COM_TI", "N")
    Call objSped.Reg_Campo_add(22, "IND_ROY_REC", "N")
    Call objSped.Reg_Campo_add(23, "IND_ROY_PAG", "N")
    Call objSped.Reg_Campo_add(24, "IND_REND_SERV", "N")
    Call objSped.Reg_Campo_add(25, "IND_PGTO_REM", "N")
    Call objSped.Reg_Campo_add(26, "IND_INOV_TEC", "N")
    Call objSped.Reg_Campo_add(27, "IND_CAP_INF", "N")
    Call objSped.Reg_Campo_add(28, "IND_PJ_HAB", "N")
    Call objSped.Reg_Campo_add(29, "IND_POLO_AM", "N")
    Call objSped.Reg_Campo_add(30, "IND_ZON_EXP", "N")
    Call objSped.Reg_Campo_add(31, "IND_AREA_COM", "N")
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0020 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0020:

   Processa_Registro_0020 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213914)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0030(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0030: DADOS CADASTRAIS
Dim lErro As Long
Dim objCidades As New ClassCidades
Dim sTelefone As String

On Error GoTo Erro_Processa_Registro_0030

    '******** registro 0030 *****************
    objSped.objSpedFiscal.sTipo = "0030"
    objSped.objSpedFiscal.lSeqPai = 0
    
    objCidades.sDescricao = objSped.objFilialEmpresa.objEndereco.sCidade
        
    If Len(Trim(objCidades.sDescricao)) > 0 Then
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Call Formata_String_Numero(objSped.objFilialEmpresa.objEndereco.sTelefone1, sTelefone)
       
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0030")
    Call objSped.Reg_Campo_add(2, "COD_NAT", "2062") '2062|Sociedade Empresária Limitada
    Call objSped.Reg_Campo_add(3, "CNAE_FISCAL", SFF_Texto(objSped.objFilialEmpresa.sCNAE, 7))
    Call objSped.Reg_Campo_add(4, "ENDERECO", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sLogradouro, 150))
    Call objSped.Reg_Campo_add(5, "NUM", SFF_Texto(objSped.objFilialEmpresa.objEndereco.lNumero, 6))
    Call objSped.Reg_Campo_add(6, "COMPL", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sComplemento, 50))
    Call objSped.Reg_Campo_add(7, "BAIRRO", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sBairro, 50))
    Call objSped.Reg_Campo_add(8, "UF", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sSiglaEstado, 2))
    Call objSped.Reg_Campo_add(9, "COD_MUN", SFF_Texto(objCidades.sCodIBGE, 7))
    Call objSped.Reg_Campo_add(10, "CEP", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sCEP, 8))
    Call objSped.Reg_Campo_add(11, "NUM_TEL", SFF_Texto(sTelefone, 15))
    Call objSped.Reg_Campo_add(12, "EMAIL", SFF_Texto(objSped.objFilialEmpresa.objEndereco.sEmail, 115))
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0030 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0030:

   Processa_Registro_0030 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213915)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0990(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_0990

    '******************* registro 0990 **********************************
    objSped.objSpedFiscal.sTipo = "0990"
    objSped.objSpedFiscal.lSeqPai = 0

    lContador = 1 'ele mesmo

    For Each objSpedRegistro In objSped.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "0" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "0990")
    Call objSped.Reg_Campo_add(2, "QTD_LIN", CStr(lContador))
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0990:

   Processa_Registro_0990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213916)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_J001(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO J001: ABERTURA DO BLOCO J
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_J001

    '******** registro J001 *****************
    objSped.objSpedFiscal.sTipo = "J001"
    objSped.objSpedFiscal.lSeqPai = 0
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "J001")
    Call objSped.Reg_Campo_add(2, "IND_MOV", "0")
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_J001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_J001:

   Processa_Registro_J001 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213917)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_J050(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO J050: PLANO DE CONTAS DO CONTRIBUINTE
Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long, alComandoJ051(1 To 6) As Long
Dim sConta As String, sContaPai As String
Dim tPlanoConta As typePlanoConta, sContaTela As String
Dim bTrata As Boolean, iAux As Integer, dtData As Date

On Error GoTo Erro_Processa_Registro_J050

    '******** registro J050 *****************
    objSped.objSpedFiscal.sTipo = "J050"
    objSped.objSpedFiscal.lSeqPai = 0

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objSped.ProxComando
    Next
    
    For iIndice = LBound(alComandoJ051) To UBound(alComandoJ051)
        alComandoJ051(iIndice) = objSped.ProxComando
    Next

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Conta, TipoConta, NivelConta, DescConta, NaturezaSped, Ativo FROM PlanoConta ORDER BY Conta", _
    tPlanoConta.sConta, tPlanoConta.iTipoConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iNaturezaSped, tPlanoConta.iAtivo)
    If lErro <> AD_SQL_SUCESSO Then gError 213918

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213919
    
    Do While lErro = AD_SQL_SUCESSO
    
        bTrata = True
        If tPlanoConta.iAtivo <> CONTA_ATIVA Then
            'Inativo
            'Verifica se tem lançamentos para própria ou filhos, se tiver não pode deixar de colocar no arquivo
            lErro = Comando_Executar(alComando(2), "SELECT FilialEmpresa FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta LIKE ? AND Data BETWEEN ? AND ? ", iAux, objSped.iFilialEmpresa, tPlanoConta.sConta & "%", objSped.dtDataIni, objSped.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 213920
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213921
            
            If lErro = AD_SQL_SEM_DADOS Then
                'Sem movimento no período
                bTrata = False
            End If
            
        End If
        
        If bTrata Then
    
            lErro = Comando_Executar(alComando(3), "SELECT DataAtualizacao FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                dtData, tPlanoConta.sConta, objSped.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 213922
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213923
        
            If lErro = AD_SQL_SEM_DADOS Then dtData = StrParaDate("01/01/2009")
            
            lErro = Mascara_MascararConta(tPlanoConta.sConta, sContaTela)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If tPlanoConta.iNivelConta > 1 Then
                
                lErro = Mascara_RetornaContaPai(tPlanoConta.sConta, sConta)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = Mascara_MascararConta(sConta, sContaPai)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Else
                sContaPai = ""
            End If
            
            objSped.objSpedFiscal.sTipo = "J050"
            objSped.objSpedFiscal.lSeqPai = 0
            
            Call objSped.Inicia_Reg
            Call objSped.Reg_Campo_add(1, "REG", "J050")
            Call objSped.Reg_Campo_add(2, "DT_ALT", SFF_Data(dtData))
            Call objSped.Reg_Campo_add(3, "COD_NAT", Format(tPlanoConta.iNaturezaSped, "00"))
            Call objSped.Reg_Campo_add(4, "IND_CTA", IIf(tPlanoConta.iTipoConta = 1, "S", "A"))
            Call objSped.Reg_Campo_add(5, "NIVEL", CStr(tPlanoConta.iNivelConta))
            Call objSped.Reg_Campo_add(6, "COD_CTA", sContaTela)
            Call objSped.Reg_Campo_add(7, "COD_CTA_SUP", sContaPai)
            Call objSped.Reg_Campo_add(8, "CTA", SFF_Texto(tPlanoConta.sDescConta))
            Call objSped.Finaliza_Reg
            
            lErro = Registro_Insere(objSped)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'se for uma conta analitica
            If tPlanoConta.iTipoConta = 3 Then
            
                lErro = Processa_Registro_J051(objSped, tPlanoConta.sConta, objSped.objSpedFiscal.lSeq, alComandoJ051)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
            
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213924
    
    Loop
    
    Processa_Registro_J050 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_J050:

   Processa_Registro_J050 = gErr

   Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213918, 213919, 213924
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)
        
        Case 213920, 213921
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case 213922, 213923
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213925)
     
    End Select

    Exit Function
    
End Function

Private Function Processa_Registro_J051(ByVal objSped As ClassSpedECFSel, ByVal sConta As String, ByVal lSeqPai As Long, alComando() As Long) As Long
'REGISTRO J051: PLANO DE CONTAS REFERENCIAL
Dim lErro As Long
Dim bSemCcl As Boolean, sCcl As String, sCclMascarado As String
Dim sContaRef  As String, sContaTelaRef As String, sCclIni As String, sCclFim As String
Dim colCtaRefAdd As New Collection, objContaCcl As ClassContaCcl, bDuplicado As Boolean
Dim objExercicio As New ClassExercicio, bIgnoraTesteCcl As Boolean
Dim dSaldoIniCta As Double, dSaldoIniCtaCcl As Double

On Error GoTo Erro_Processa_Registro_J051

    '******** registro J051 *****************
    objSped.objSpedFiscal.sTipo = "J051"
    objSped.objSpedFiscal.lSeqPai = lSeqPai
   
    lErro = CF("Exercicio_Le_Data", objSped.dtDataIni, objExercicio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    bSemCcl = True
'
'    sCcl = String(STRING_CCL, 0)
'
'    lErro = Comando_Executar(alComando(1), "SELECT Ccl FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl <> ? AND Conta = ? AND Data BETWEEN ? AND ?  ", _
'    sCcl, objSped.iFilialEmpresa, "", sConta, objSped.dtDataIni, objSped.dtDataFim)
'    If lErro <> AD_SQL_SUCESSO Then gError 213926
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213927
'
'    If lErro = AD_SQL_SUCESSO Then bSemCcl = False ' Essa conta trabalha com centro de custo
        
    bIgnoraTesteCcl = False
    If UCase(left(gsNomeEmpresa, 5)) = "INPAL" And (left(sConta, 1) = "1" Or left(sConta, 1) = "2") Then
        bSemCcl = True
        bIgnoraTesteCcl = True
    End If
    
    If Not bIgnoraTesteCcl Then
        
        sCcl = String(STRING_CCL, 0)
        
        lErro = Comando_Executar(alComando(1), "SELECT Ccl FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl <> ? AND Conta = ? AND Data BETWEEN ? AND ?  ", _
        sCcl, iFilialEmpresa, "", sConta, objExercicio.dtDataInicio, objExercicio.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203149
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
        
        If lErro = AD_SQL_SUCESSO Then
        
            bSemCcl = False ' Essa conta trabalha com centro de custo
        
            'Se achou lançamento com ccl procura sem ccl para mesma conta/período
            lErro = Comando_Executar(alComando(4), "SELECT Ccl FROM Lancamentos WHERE Ccl = ? AND Conta = ? AND (FilialEmpresa = ? OR ?=?) AND Data BETWEEN ? AND ? AND ABS(Valor) > 0.001  AND (Aglutinado = 0 OR Aglutinado = 1) ", _
            sCcl, "", sConta, objSped.iFilialEmpresa, objSped.iFilialEmpresa, EMPRESA_TODA, objExercicio.dtDataInicio, objExercicio.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 203149
    
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150

            If lErro = AD_SQL_SUCESSO Then
                bSemCcl = True 'Não trabalha integralmente
            Else
                'Testa o saldo inicial com CCL e total, se não forem iguais não pode considerar que a conta trabalha com ccl
                lErro = Comando_Executar(alComando(5), "SELECT SldIni FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? ", _
                dSaldoIniCta, objSped.iFilialEmpresa, objExercicio.iExercicio, sConta)
                If lErro <> AD_SQL_SUCESSO Then gError 203149
        
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                
                lErro = Comando_Executar(alComando(6), "SELECT SUM(SldIni) FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? ", _
                dSaldoIniCtaCcl, objSped.iFilialEmpresa, objExercicio.iExercicio, sConta)
                If lErro <> AD_SQL_SUCESSO Then gError 203149
        
                lErro = Comando_BuscarPrimeiro(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203150
                
                If Abs(dSaldoIniCta - dSaldoIniCtaCcl) > DELTA_VALORMONETARIO Then
                    bSemCcl = True 'Não trabalha integralmente
                End If
            
            End If
            
        Else
            bSemCcl = True
        End If
        
    End If
    
    sContaRef = String(STRING_CONTA, 0)
    sContaTelaRef = String(STRING_CONTA, 0)
    sCclIni = String(STRING_CCL, 0)
    sCclFim = String(STRING_CCL, 0)

    lErro = Comando_Executar(alComando(2), "SELECT P.ContaImp, PC.Conta, PC.CclInicial, PC.CclFinal FROM PlanoContaRefConfig AS PC, PlanoContaRef As P WHERE PC.CodigoModelo = P.CodigoModelo AND PC.Conta = P.Conta AND PC.CodigoModelo = ? AND ContaInicial <= ? And ContaFinal >= ? ", _
    sContaTelaRef, sContaRef, sCclIni, sCclFim, objSped.lModelo, sConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 213928

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213929

    Do While lErro = AD_SQL_SUCESSO
    
        'Inclui a associação por ccl se o ccl estiver associado a conta
        If Not bSemCcl Then
            'se o centro de custo da conta  estiver preenchido
        
            sCcl = String(STRING_CCL, 0)
            
            lErro = Comando_Executar(alComando(3), "SELECT DISTINCT Ccl FROM MvPerCcl WHERE FilialEmpresa = ? AND (Ccl >= ? OR ? = ?) AND (Ccl <= ? OR ? = ?) AND Conta = ? AND Exercicio = ? ORDER BY Ccl", _
            sCcl, iFilialEmpresa, sCclIni, sCclIni, "", sCclFim, sCclFim, "", sConta, objExercicio.iExercicio)
            'lErro = Comando_Executar(alComando(3), "SELECT DISTINCT Ccl FROM MvDiaCcl WHERE FilialEmpresa = ? AND (Ccl >= ? OR ? = ?) AND (Ccl <= ? OR ? = ?) AND Conta = ? AND Data BETWEEN ? AND ? ORDER BY Ccl", _
            'sCcl, iFilialEmpresa, sCclIni, sCclIni, "", sCclFim, sCclFim, "", sConta, objSped.dtDataIni, objSped.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 213930
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213931
            
            Do While lErro = AD_SQL_SUCESSO
                                 
                lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                bDuplicado = False
                For Each objContaCcl In colCtaRefAdd
                    If SFF_Texto(sCclMascarado) = objContaCcl.sCcl And SFF_Texto(sContaTelaRef) = objContaCcl.sConta Then
                        bDuplicado = True
                        Exit For
                    End If
                Next
                                
                If Not bDuplicado Then
                
                    Set objContaCcl = New ClassContaCcl
                    objContaCcl.sCcl = SFF_Texto(sCclMascarado)
                    objContaCcl.sConta = SFF_Texto(sContaTelaRef)
                    colCtaRefAdd.Add objContaCcl
                
                    Call objSped.Inicia_Reg
                    Call objSped.Reg_Campo_add(1, "REG", "J051")
                    Call objSped.Reg_Campo_add(2, "COD_CCUS", SFF_Texto(sCclMascarado))
                    Call objSped.Reg_Campo_add(3, "COD_CTA_REF", SFF_Texto(sContaTelaRef))
                    Call objSped.Finaliza_Reg
                    
                    lErro = Registro_Insere(objSped)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    
                End If
                                
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213932
                
            Loop
            
        Else
        
'            bDuplicado = False
'            For Each objContaCcl In colCtaRefAdd
'                If SFF_Texto(sCclMascarado) = objContaCcl.sCcl And SFF_Texto(sContaTelaRef) = objContaCcl.sConta Then
'                    bDuplicado = True
'                    Exit For
'                End If
'            Next
'
'            If Not bDuplicado Then
            
                Set objContaCcl = New ClassContaCcl
                objContaCcl.sCcl = ""
                objContaCcl.sConta = SFF_Texto(sContaTelaRef)
                colCtaRefAdd.Add objContaCcl
                
                Call objSped.Inicia_Reg
                Call objSped.Reg_Campo_add(1, "REG", "J051")
                Call objSped.Reg_Campo_add(2, "COD_CCUS", "")
                Call objSped.Reg_Campo_add(3, "COD_CTA_REF", SFF_Texto(sContaTelaRef))
                Call objSped.Finaliza_Reg
            
                lErro = Registro_Insere(objSped)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
'            End If

                Exit Do
                'quando não tem ccl só pode ter um J051 para cada J050
                'algumas contas sem movimento da Inpal estavam mapeadas em mais de uma conta referencial (ccl diferentes)
                'se a conta tivesse usado ccl cairia na 1a clausula do IF, como caiu aqui vai pegar uma delas só para não dar erro e descartar a/as outras
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213933
        
    Loop
    
    Processa_Registro_J051 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_J051:

   Processa_Registro_J051 = gErr

   Select Case gErr

        Case 213926, 213927
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", gErr)

        Case 213928, 213929, 213933
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 213930 To 213932
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213934)
     
    End Select

    Exit Function
    
End Function

Private Function Processa_Registro_J100(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO J100: CENTRO DE CUSTOS
Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long, bTrata As Boolean
Dim tCcl As typeCcl, sCclTela As String
Dim iAux As Integer, dtData As Date

On Error GoTo Erro_Processa_Registro_J100

    '******** registro J100 *****************
    objSped.objSpedFiscal.sTipo = "J100"
    objSped.objSpedFiscal.lSeqPai = 0

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objSped.ProxComando
    Next
    
    tCcl.sCcl = String(STRING_CCL, 0)
    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Ccl, DescCcl, AtivoCcl FROM Ccl ORDER BY Ccl", tCcl.sCcl, tCcl.sDescCcl, tCcl.iAtivo)
    If lErro <> AD_SQL_SUCESSO Then gError 213935

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213936
    
    Do While lErro = AD_SQL_SUCESSO
    
        bTrata = True
        If tCcl.iAtivo <> MARCADO Then
            'Inativo
            'Verifica se tem lançamentos para própria ou filhos, se tiver não pode deixar de colocar no arquivo
            lErro = Comando_Executar(alComando(2), "SELECT FilialEmpresa FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl LIKE ? AND Data BETWEEN ? AND ? ", iAux, objSped.iFilialEmpresa, tCcl.sCcl & "%", objSped.dtDataIni, objSped.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 213937
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213938
            
            If lErro = AD_SQL_SEM_DADOS Then
                'Sem movimento no período
                bTrata = False
            End If
            
        End If
        
        If bTrata Then
    
            lErro = Comando_Executar(alComando(3), "SELECT DataAtualizacao FROM CclHistorico WHERE CodigoCcl = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC", _
                dtData, tCcl.sCcl, objSped.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 213939
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213940
        
            If lErro = AD_SQL_SEM_DADOS Then dtData = StrParaDate("01/01/2009")
            
            lErro = Mascara_MascararCcl(tCcl.sCcl, sCclTela)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Call objSped.Inicia_Reg
            Call objSped.Reg_Campo_add(1, "REG", "J100")
            Call objSped.Reg_Campo_add(2, "DT_ALT", SFF_Data(dtData))
            Call objSped.Reg_Campo_add(3, "COD_CCUS", sCclTela)
            Call objSped.Reg_Campo_add(4, "CCUS", SFF_Texto(tCcl.sDescCcl))
            Call objSped.Finaliza_Reg
            
            lErro = Registro_Insere(objSped)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213941
    
    Loop
    
    Processa_Registro_J100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_J100:

   Processa_Registro_J100 = gErr

   Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213935, 213936, 213941
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)
        
        Case 213937, 213938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
        
        Case 213939, 213940
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213942)
     
    End Select

    Exit Function
    
End Function

Private Function Processa_Registro_J990(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO J990: ENCERRAMENTO DO BLOCO J
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_J990

    '******************* registro J990 **********************************
    objSped.objSpedFiscal.sTipo = "J990"
    objSped.objSpedFiscal.lSeqPai = 0

    lContador = 1 'ele mesmo

    For Each objSpedRegistro In objSped.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "J" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "J990")
    Call objSped.Reg_Campo_add(2, "QTD_LIN", CStr(lContador))
    Call objSped.Finaliza_Reg
    
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_J990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_J990:

   Processa_Registro_J990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213943)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9001(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 9001: ABERTURA DO BLOCO 9
'Este registro deve sempre ser gerado e representa a abertura do Bloco 9.
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_9001

    '******************* registro 9001 **********************************
    objSped.objSpedFiscal.sTipo = "9001"
    objSped.objSpedFiscal.lSeqPai = 0

    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "9001")
    Call objSped.Reg_Campo_add(2, "IND_DAD", "0")
    Call objSped.Finaliza_Reg
        
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9001:

   Processa_Registro_9001 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213944)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9900(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 9900: REGISTROS DO ARQUIVO
'Este registro informa a quantidade total de registros por tipo de registro do arquivo da ECF.
'Devem ser mostradas a quantidade de ocorrências de todos os registros que aparecem no arquivo da ECF,
'inclusive do próprio registro 9900, que será preenchido por último.
'Além disso, deve ser informado o código da versão da tabela dinâmica utilizada (atualmente a versão é 1.00) e a
'identificação da tabela dinâmica utilizada (será preenchido pelo sistema).
'Ambos os campos (9900.VERSAO e 9900.ID_TAB_DIN), atualmente, são não obrigatórios.
Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9900

    '******************* registro 9900 **********************************
    objSped.objSpedFiscal.sTipo = "9900"
    objSped.objSpedFiscal.lSeqPai = 0

    Call Registro_Adiciona_Count("9990", objSped.colSpedRegistro, 1)
    Call Registro_Adiciona_Count("9999", objSped.colSpedRegistro, 1)
    Call Registro_Adiciona_Count(objSped.objSpedFiscal.sTipo, objSped.colSpedRegistro, objSped.colSpedRegistro.Count + 1)

    For Each objSpedRegistro In objSped.colSpedRegistro

        Call objSped.Inicia_Reg
        Call objSped.Reg_Campo_add(1, "REG", "9900")
        Call objSped.Reg_Campo_add(2, "REG_BLC", objSpedRegistro.sTipo)
        Call objSped.Reg_Campo_add(3, "QTD_REG_BLC", CStr(objSpedRegistro.lCount))
        Call objSped.Reg_Campo_add(4, "VERSAO", "")
        Call objSped.Reg_Campo_add(5, "ID_TAB_DIN", "")
        Call objSped.Finaliza_Reg

        lErro = Registro_Insere(objSped)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next

    Processa_Registro_9900 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9900:

   Processa_Registro_9900 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213945)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9990(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
'Este registro informa a quantidade total de registros do Bloco 9.
'Há que se ressaltar que vários registros ocorrem mais de uma vez.
'Portanto, na contagem dos registros, deve-se levar em consideração todos os registros informados, ainda que repetidos.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9990

    '******************* registro 9990 **********************************
    objSped.objSpedFiscal.sTipo = "9990"
    objSped.objSpedFiscal.lSeqPai = 0

    For Each objSpedRegistro In objSped.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "9" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next

    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "9990")
    Call objSped.Reg_Campo_add(2, "QTD_LIN", CStr(lContador))
    Call objSped.Finaliza_Reg
            
    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9990:

   Processa_Registro_9990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213946)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9999(ByVal objSped As ClassSpedECFSel) As Long
'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
'Este registro informa a quantidade total de registros (linhas) do arquivo da ECF.
'Há que se ressaltar que vários registros ocorrem mais de uma vez.
'Portanto, na contagem dos registros, deve-se levar em consideração todos os registros informados, ainda que repetidos.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9999

    '******************* registro 9999 **********************************
    objSped.objSpedFiscal.sTipo = "9999"
    objSped.objSpedFiscal.lSeqPai = 0
    
    For Each objSpedRegistro In objSped.colSpedRegistro
        lContador = lContador + objSpedRegistro.lCount
    Next
    
    Call objSped.Inicia_Reg
    Call objSped.Reg_Campo_add(1, "REG", "9999")
    Call objSped.Reg_Campo_add(2, "QTD_LIN", CStr(lContador))
    Call objSped.Finaliza_Reg

    lErro = Registro_Insere(objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9999 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9999:

   Processa_Registro_9999 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213947)
    
    End Select
    
    Exit Function
    
End Function

Private Function SFF_Texto(ByVal sTexto As String, Optional ByVal lTamMax As Long = 0, Optional bFazTrim As Boolean = True, Optional bRemoveEspacoDuplo As Boolean = True, Optional bRemoveAcentos As Boolean = False, Optional bUpperCase As Boolean = False) As String
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

Private Function SFF_Data(ByVal dtData As Date) As String
    If dtData <> DATA_NULA Then
        SFF_Data = Format(dtData, "ddmmyyyy")
    Else
        SFF_Data = ""
    End If
End Function

Private Function SFF_Num(ByVal dNum As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Num = Format(dNum, sFormato)
End Function

Private Function SFF_Perc(ByVal dPerc As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Perc = Format(dPerc * 100, sFormato)
End Function
