VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISSpedPis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'http://www1.receita.fazenda.gov.br/sistemas/efd-contribuicoes/

'??? Processa_Registro_C120 só deve ser chamado para nf de importacao

Function Gera_Sped_Fiscal_Pis(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Gera_Sped_Fiscal_Pis

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209355
    
    'Obtem totalizadores, grava SpedDocFiscais e lê dados como filial empresa e outros que serão usados no processamento
    Call objEFD.objTela.SPED_Status("Obtendo dados inicias", "")
    lErro = SpedFiscal_Obtem_DadosInicias(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'BLOCO 0: Abertura, Identificação e Referências
    Call objEFD.objTela.SPED_Status("Processando Bloco 0", "")
    lErro = Processa_Bloco_0(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO A: Documentos Fiscais - Serviços (ISS)
    Call objEFD.objTela.SPED_Status("Processando Bloco A", "")
    lErro = Processa_Bloco_A(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO C: Documentos Fiscais I – Mercadorias (ICMS/IPI)
    Call objEFD.objTela.SPED_Status("Processando Bloco C", "")
    lErro = Processa_Bloco_C(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO D: Documentos Fiscais II – Serviços (ICMS)
    Call objEFD.objTela.SPED_Status("Processando Bloco D", "")
    lErro = Processa_Bloco_D(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO F: Demais Documentos e Operações
    Call objEFD.objTela.SPED_Status("Processando Bloco F", "")
    lErro = Processa_Bloco_F(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO M: Apuração da Contribuição e Crédito de PIS/PASEP e da COFINS
    Call objEFD.objTela.SPED_Status("Processando Bloco M", "")
    lErro = Processa_Bloco_M(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO P: BLOCO P: APURAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA
    Call objEFD.objTela.SPED_Status("Processando Bloco P", "")
    lErro = Processa_Bloco_P(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO 1: Complemento da Escrituração – Controle de Saldos de Créditos e de Retenções, Operações Extemporâneas e Outras Informações
    Call objEFD.objTela.SPED_Status("Processando Bloco 1", "")
    lErro = Processa_Bloco_1(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'BLOCO 9: Controle e Encerramento do Arquivo Digital
    Call objEFD.objTela.SPED_Status("Processando Bloco 9", "")
    lErro = Processa_Bloco_9(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Gera o arquivo texto a partir dos dados da tabela
    lErro = SpedFiscal_Gravar_Arq(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209356
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_SPEDFISCALPIS_SUCESSO", objEFD.sDiretorio)
    
    Gera_Sped_Fiscal_Pis = SUCESSO
    
    Exit Function
    
Erro_Gera_Sped_Fiscal_Pis:

    Gera_Sped_Fiscal_Pis = gErr
    
    Select Case gErr
    
        Case 209355
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 209356
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209357)
     
    End Select
   
    Call Transacao_Rollback
   
    Exit Function
    
End Function

Private Function Processa_Bloco_0(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Bloco_0

    '========================================================================================
    '=============================== INÍCIO DO BLOCO 0 ======================================
    '================== BLOCO 0: ABERTURA, IDENTIFICAÇÃO E REFERÊNCIAS. =====================
    '========================================================================================

    'REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA
    Call objEFD.objTela.SPED_Status("Processando Registro 0000", "")
    lErro = Processa_Registro_0000(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0001: ABERTURA DO BLOCO 0
    Call objEFD.objTela.SPED_Status("Processando Registro 0001", "")
    lErro = Processa_Registro_0001(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'REGISTRO 0100: DADOS DO CONTABILISTA
    Call objEFD.objTela.SPED_Status("Processando Registro 0100", "")
    lErro = Processa_Registro_0100(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0110: REGIMES DE APURAÇÃO DA CONTRIBUIÇÃO SOCIAL E DE APROPRIAÇÃO DE CRÉDITO
    Call objEFD.objTela.SPED_Status("Processando Registro 0110", "")
    lErro = Processa_Registro_0110(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'REGISTRO 0140: TABELA DE CADASTRO DE ESTABELECIMENTO
    Call objEFD.objTela.SPED_Status("Processando Registro 0140", "")
    lErro = Processa_Registro_0140(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 0500: PLANO DE CONTAS CONTÁBEIS
    If objEFD.colModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        Call objEFD.objTela.SPED_Status("Processando Registro 0500", "")
        lErro = Processa_Registro_0500(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'REGISTRO 0600: CENTRO DE CUSTOS
        Call objEFD.objTela.SPED_Status("Processando Registro 0600", "")
        lErro = Processa_Registro_0600(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
    Call objEFD.objTela.SPED_Status("Processando Registro 0990", "")
    lErro = Processa_Registro_0990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Processa_Bloco_0 = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_0:

    Processa_Bloco_0 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209358)
     
    End Select
    
End Function

Private Function Processa_Bloco_A(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBlocoA As Boolean

On Error GoTo Erro_Processa_Bloco_A

    '========================================================================================
    '=============================== INÍCIO DO BLOCO A ======================================
    '=========== BLOCO A: DOCUMENTOS FISCAIS - SERVIÇOS (NÃO SUJEITOS AO ICMS) ==============
    '========================================================================================
    
    'REGISTRO A001: ABERTURA DO BLOCO A
    Call objEFD.objTela.SPED_Status("Processando Registro A001", "")
    lErro = Processa_Registro_A001(objEFD, bAbreBlocoA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If bAbreBlocoA Then
    
        'REGISTRO A010: IDENTIFICAÇÃO DO ESTABELECIMENTO
        'REGISTRO A100: DOCUMENTO - NOTA FISCAL DE SERVIÇO
        'REGISTRO A110: COMPLEMENTO DO DOCUMENTO - INFORMAÇÃO COMPLEMENTAR DA NF
        'REGISTRO A111: PROCESSO REFERENCIADO
        'REGISTRO A120: INFORMAÇÃO COMPLEMENTAR - OPERAÇÕES DE IMPORTAÇÃO
        'REGISTRO A170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco A", "")
        lErro = Processa_Registro_A010(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    'REGISTRO A990: ENCERRAMENTO DO BLOCO A
    Call objEFD.objTela.SPED_Status("Processando Registro A990", "")
    lErro = Processa_Registro_A990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_A = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_A:

    Processa_Bloco_A = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209359)
     
    End Select
    
End Function

Private Function Processa_Bloco_C(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBlocoC As Boolean

On Error GoTo Erro_Processa_Bloco_C

    '========================================================================================
    '=============================== INÍCIO DO BLOCO C ======================================
    '=============== BLOCO C: DOCUMENTOS FISCAIS I - MERCADORIAS (ICMS/IPI) =================
    '========================================================================================
    
    'REGISTRO C001: ABERTURA DO BLOCO C
    Call objEFD.objTela.SPED_Status("Processando Registro C001", "")
    lErro = Processa_Registro_C001(objEFD, bAbreBlocoC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBlocoC Then

        'REGISTRO C010: IDENTIFICAÇÃO DO ESTABELECIMENTO
        'REGISTRO C100: DOCUMENTO - NOTA FISCAL (CÓDIGO 01), NOTA FISCAL AVULSA (CÓDIGO 1B), NOTA FISCAL DE PRODUTOR (CÓDIGO 04) e NF-e (CÓDIGO 55)
        'REGISTRO C110: COMPLEMENTO DO DOCUMENTO - INFORMAÇÃO COMPLEMENTAR DA NOTA FISCAL (CÓDIGOS 01, 1B, 04 e 55)
        'REGISTRO C111: PROCESSO REFERENCIADO
        'REGISTRO C120: COMPLEMENTO DO DOCUMENTO - OPERAÇÕES DE IMPORTAÇÃO (CÓDIGO 01)
        'REGISTRO C170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO (CÓDIGOS 01, 1B, 04 e 55)
        'REGISTRO C500: NOTA FISCAL/CONTA DE ENERGIA ELÉTRICA (CÓDIGO 06), NOTA FISCAL/CONTA DE FORNECIMENTO D'ÁGUA CANALIZADA (CÓDIGO 29) E NOTA FISCAL CONSUMO FORNECIMENTO DE GÁS (CÓDIGO 28) - DOCUMENTOS DE ENTRADA/AQUISIÇÃO COM CRÉDITO
        'REGISTRO C501: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 06, 28 e 29) - PIS/PASEP
        'REGISTRO C505: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 06, 28 e 29) - COFINS
        'REGISTRO C509: PROCESSO REFERENCIADO
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco C", "")
        lErro = Processa_Registro_C010(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO C990: ENCERRAMENTO DO BLOCO C
    Call objEFD.objTela.SPED_Status("Processando Registro C990", "")
    lErro = Processa_Registro_C990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_C = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_C:

    Processa_Bloco_C = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209360)
     
    End Select
    
End Function

Private Function Processa_Bloco_D(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBloco As Boolean

On Error GoTo Erro_Processa_Bloco_D

    '========================================================================================
    '=============================== INÍCIO DO BLOCO D ======================================
    '================== BLOCO D: DOCUMENTOS FISCAIS II - SERVIÇOS (ICMS) ====================
    '========================================================================================
    
    'REGISTRO D001: ABERTURA DO BLOCO D
    Call objEFD.objTela.SPED_Status("Processando Registro D001", "")
    lErro = Processa_Registro_D001(objEFD, bAbreBloco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBloco Then

        'REGISTRO D010: IDENTIFICAÇÃO DO ESTABELECIMENTO
        'REGISTRO D100: Aquisição de Serviços de Transportes (Códigos 07, 08, 8B, 09, 10, 11, 26, 27 e 57)
        'REGISTRO D101: Complemento do Documento de Transporte – PIS/PASEP
        'REGISTRO D105: Complemento do Documento de Transporte – COFINS
        'REGISTRO D111: Processo Referenciado
        'REGISTRO D500: Nota Fiscal de Serviço de Comunicação (Código 21) e Serviço de Telecomunicação (Código 22) - Operação de Aquisição com Direito a Crédito
        'REGISTRO D501: Complemento da Operação (Código 21 e 22) – PIS/PASEP
        'REGISTRO D505: Complemento da Operação (Código 21 e 22) – COFINS
        'REGISTRO D509: PROCESSO REFERENCIADO
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco D", "")
        lErro = Processa_Registro_D010(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO D990: ENCERRAMENTO DO BLOCO D
    Call objEFD.objTela.SPED_Status("Processando Registro D990", "")
    lErro = Processa_Registro_D990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Processa_Bloco_D = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_D:

    Processa_Bloco_D = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209361)
     
    End Select
    
End Function

Private Function Processa_Bloco_F(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBloco As Boolean

On Error GoTo Erro_Processa_Bloco_F

    '========================================================================================
    '=============================== INÍCIO DO BLOCO F ======================================
    '====================== BLOCO F: DEMAIS DOCUMENTOS E OPERAÇÕES ==========================
    '========================================================================================
    
    'REGISTRO F001: ABERTURA DO BLOCO F
    Call objEFD.objTela.SPED_Status("Processando Registro F001", "")
    lErro = Processa_Registro_F001(objEFD, bAbreBloco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBloco Then
        
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco F", "")
        lErro = Processa_Registro_F010(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO F990: ENCERRAMENTO DO BLOCO F
    Call objEFD.objTela.SPED_Status("Processando Registro F990", "")
    lErro = Processa_Registro_F990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_F = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_F:

    Processa_Bloco_F = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209362)
     
    End Select
    
End Function

Private Function Processa_Bloco_M(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBloco As Boolean

On Error GoTo Erro_Processa_Bloco_M

    '========================================================================================
    '=============================== INÍCIO DO BLOCO M ======================================
    '======= BLOCO M: Apuração da Contribuição e Crédito de PIS/PASEP e da COFINS ===========
    '========================================================================================
    
    'REGISTRO M001: ABERTURA DO BLOCO M
    Call objEFD.objTela.SPED_Status("Processando Registro M001", "")
    lErro = Processa_Registro_M001(objEFD, bAbreBloco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBloco Then
        
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco M", "")
        
        lErro = Processa_Registro_M200(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Processa_Registro_M600(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO M990: ENCERRAMENTO DO BLOCO M
    Call objEFD.objTela.SPED_Status("Processando Registro M990", "")
    lErro = Processa_Registro_M990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_M = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_M:

    Processa_Bloco_M = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209363)
     
    End Select
    
End Function

Private Function Processa_Bloco_P(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBloco As Boolean
Dim dTotalContrPrev As Double

On Error GoTo Erro_Processa_Bloco_P

    '========================================================================================
    '=============================== INÍCIO DO BLOCO P ======================================
    '====== BLOCO P: APURAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA ==========
    '========================================================================================
    
    'REGISTRO P001: ABERTURA DO BLOCO P
    Call objEFD.objTela.SPED_Status("Processando Registro P001", "")
    lErro = Processa_Registro_P001(objEFD, bAbreBloco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBloco Then
        
        'REGISTRO P010: IDENTIFICAÇÃO DO ESTABELECIMENTO
        'REGISTRO P100: CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA
        Call objEFD.objTela.SPED_Status("Processando Registros do bloco P", "")
        lErro = Processa_Registro_P010(objEFD, dTotalContrPrev)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'REGISTRO P200: CONSOLIDAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA
        Call objEFD.objTela.SPED_Status("Processando Registros P200", "")
        lErro = Processa_Registro_P200(objEFD, dTotalContrPrev)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO P990: ENCERRAMENTO DO BLOCO P
    Call objEFD.objTela.SPED_Status("Processando Registro P990", "")
    lErro = Processa_Registro_P990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_P = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_P:

    Processa_Bloco_P = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209364)
     
    End Select
    
End Function

Private Function Processa_Bloco_1(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim bAbreBloco As Boolean

On Error GoTo Erro_Processa_Bloco_1

    '========================================================================================
    '=============================== INÍCIO DO BLOCO 1 ======================================
    '==== BLOCO 1: COMPLEMENTO DA ESCRITURAÇÃO – CONTROLE DE SALDOS DE CRÉDITOS E DE ========
    '============== RETENÇÕES, OPERAÇÕES EXTEMPORÂNEAS E OUTRAS INFORMAÇÕES =================
    '========================================================================================
    
    'REGISTRO 1001: ABERTURA DO BLOCO 1
    Call objEFD.objTela.SPED_Status("Processando Registro 1001", "")
    lErro = Processa_Registro_1001(objEFD, bAbreBloco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If bAbreBloco Then
        
        'REGISTRO 1100: CONTROLE DE CRÉDITOS FISCAIS – PIS/PASEP
        Call objEFD.objTela.SPED_Status("Processando Registro 1100", "")
        lErro = Processa_Registro_1100(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 1500: CONTROLE DE CRÉDITOS FISCAIS – COFINS
        Call objEFD.objTela.SPED_Status("Processando Registro 1500", "")
        lErro = Processa_Registro_1500(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 1900: CONSOLIDAÇÃO DOS DOCUMENTOS EMITIDOS NO PERÍODO POR
        'PESSOA JURÍDICA SUBMETIDA AO REGIME DE TRIBUTAÇÃO COM BASE NO
        'LUCRO PRESUMIDO – REGIME DE CAIXA OU DE COMPETÊNCIA
        Call objEFD.objTela.SPED_Status("Processando Registro 1900", "")
        lErro = Processa_Registro_1900(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    'REGISTRO 1990: ENCERRAMENTO DO BLOCO 1
    Call objEFD.objTela.SPED_Status("Processando Registro 1990", "")
    lErro = Processa_Registro_1990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_1 = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_1:

    Processa_Bloco_1 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209364)
     
    End Select
    
End Function

Private Function Processa_Bloco_9(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Bloco_9

    '========================================================================================
    '=============================== INÍCIO DO BLOCO 9 ======================================
    '================ BLOCO 9: CONTROLE E ENCERRAMENTO DO ARQUIVO DIGITAL ===================
    '========================================================================================
    
    'REGISTRO 9001: ABERTURA DO BLOCO 9
    Call objEFD.objTela.SPED_Status("Processando Registro 9001", "")
    lErro = Processa_Registro_9001(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9900: REGISTROS DO ARQUIVO
    Call objEFD.objTela.SPED_Status("Processando Registro 9900", "")
    lErro = Processa_Registro_9900(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
    Call objEFD.objTela.SPED_Status("Processando Registro 9990", "")
    lErro = Processa_Registro_9990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
    Call objEFD.objTela.SPED_Status("Processando Registro 9999", "")
    lErro = Processa_Registro_9999(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_9 = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_9:

    Processa_Bloco_9 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209365)
     
    End Select
    
End Function

Private Function SpedFiscal_Obtem_DadosInicias(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim lNumIntSped As Long
Dim iCodVersao As Integer
Dim lTransacao As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais, lContador As Long

On Error GoTo Erro_SpedFiscal_Obtem_DadosInicias

    'Abertura de Comando
    For iIndice = 0 To objEFD.NUM_MAX_COMANDOS
        objEFD.alComando(iIndice) = Comando_Abrir()
        If objEFD.alComando(iIndice) = 0 Then gError 209366
    Next
    
    'Obtem o total de registros que vão atualizar a barra de progressão
    lErro = Registros_Totaliza(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call objEFD.objTela.Inicia_Processo(objEFD.lNumTotalPrevisto)
        
    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_SPEDFISCALPIS", lNumIntSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Obtém a versão do layout
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT Codigo From SpedFisPisVersaoLeiaute WHERE DataInicio <= ? ORDER BY DataInicio DESC ", iCodVersao, objEFD.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 209367
    
    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209368
    
    'versao nao cadastrada
    If lErro <> AD_SQL_SUCESSO Then gError 209369
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa 'FILIAL_MATRIZ 'iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    objEFD.lNumIntSped = lNumIntSped
    objEFD.iCodVersao = iCodVersao
    
    objEFD.objSpedFiscal.iFilialEmpresa = objEFD.iFilialEmpresa
    objEFD.objSpedFiscal.lNumIntSped = objEFD.lNumIntSped
    
    Set objEFD.objFilialEmpresa = objFilialEmpresa
    
    If objEFD.iCodVersao < 3 Or objEFD.objFilialEmpresa.iLucroPresumido <> MARCADO Or objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = MARCADO Or objEFD.objFilialEmpresa.iPisCofinsRegCumTipo <> 1 Then
        objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA
    Else
        objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA
    End If
    
    'Grava SpedFiscalDocs e SpedFiscalDocItens
    lErro = SpedFiscal_Grava_Tabelas(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    SpedFiscal_Obtem_DadosInicias = SUCESSO

    Exit Function
    
Erro_SpedFiscal_Obtem_DadosInicias:

   SpedFiscal_Obtem_DadosInicias = gErr

   Select Case gErr
   
        Case 209366
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209367, 209368
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISPISVERSAOLEIAUTE", gErr)
            
        Case 209369
            Call Rotina_Erro(vbOKOnly, "ERRO_EFD_PIS_COFINS_SEM_VERSAO", gErr, Format(objEFD.dtDataFim, "dd/mm/yyyy"))
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209370)
     
    End Select
    
End Function

Private Function Registros_Totaliza(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'Verifica os registros com grande processamento para ajustar a barra de progressão

Dim lErro As Long, iAno As Integer
Dim lContadorItensNF As Long
Dim sSQLNF As String

On Error GoTo Erro_Registros_Totaliza
    
    Call objEFD.objTela.SPED_Status("Totalizando Registros", "NFiscal")
    
    sSQLNF = "SELECT COUNT (*) "
    sSQLNF = sSQLNF & "FROM NFiscal AS N, TiposDocInfo AS T, TributacaoDoc AS TD, ItensNFiscal AS I "
    sSQLNF = sSQLNF & "WHERE I.NumIntNF = N.NumIntDoc AND N.TipoNFiscal = T.Codigo AND TD.TipoDoc = 0 AND TD.NumIntDoc = N.NumIntDoc AND "
    sSQLNF = sSQLNF & "((T.Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ( "
    sSQLNF = sSQLNF & "(T.Tipo = 1 Or T.Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND N.FilialEmpresa < 50 "
    
    lErro = Comando_Executar(objEFD.ProxComando, sSQLNF, lContadorItensNF, objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 209371

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209372

    objEFD.lNumTotalItensNF = lContadorItensNF
    
    'É uma previsão mais ou menos que vai ser ajustada no decorrer do processo
    objEFD.lNumTotalPrevisto = 3 * lContadorItensNF '(Ler, Processar e Gravar)
    
    Registros_Totaliza = SUCESSO

    Exit Function
    
Erro_Registros_Totaliza:

   Registros_Totaliza = gErr

   Select Case gErr

        Case 209371, 209372
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209373)
     
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

Private Function Registro_Insere_BD(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim objSpedFiscal As ClassSpedFiscal
Dim sAux As String, iSeq As Long, sReg As String, sErro As String

On Error GoTo Erro_Registro_Insere_BD

    sErro = "01"
    Set objSpedFiscal = objEFD.objSpedFiscal

    sErro = "02"
    If left(objSpedFiscal.sTipo, 2) <> "99" Then Call Registro_Adiciona_Count(objEFD.objSpedFiscal.sTipo, objEFD.colSpedRegistro, 1)

    sErro = "03"
    lErro = Comando_Executar(objEFD.lComandoInsert, "INSERT INTO SpedFiscalPis( NumIntSped, Seq, FilialEmpresa, Tipo,  SeqPai ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.lSeqPai)
    sErro = "04"
    If lErro <> AD_SQL_SUCESSO Then gError 209374
    
    sErro = "05"
    sAux = objSpedFiscal.sRegistro

    sErro = "06"
    iSeq = 0
    sErro = "07"
    Do While Len(Trim(sAux)) > 0
    
    sErro = "08"
        iSeq = iSeq + 1
        
    sErro = "09"
        sReg = left(sAux, STRING_MAXIMO)
    
    sErro = "10"
        If Len(Trim(sAux)) > STRING_MAXIMO Then
    sErro = "11"
            sAux = Mid(sAux, STRING_MAXIMO + 1)
        Else
    sErro = "12"
            sAux = ""
        End If
        
    sErro = "13"
        lErro = Comando_Executar(objEFD.lComandoInsertReg, "INSERT INTO SpedFiscalPisReg(NumIntSped, Seq, SeqReg, Registro ) VALUES (?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, iSeq, sReg)
    sErro = "14"
        If lErro <> AD_SQL_SUCESSO Then gError 209375
    
    sErro = "15"
    Loop
    
    sErro = "16"
    lErro = objEFD.objTela.Processa_Item(True)
    sErro = "17"
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sErro = "18"
    objEFD.lNumTotalProcessados = objEFD.lNumTotalProcessados + 1 'Total de processamentos (3 por registro)
    sErro = "19"
    objEFD.lNumTotalReal = objEFD.lNumTotalReal + 1 'Total real de registros processados
    sErro = "20"

    Registro_Insere_BD = SUCESSO

    Exit Function

Erro_Registro_Insere_BD:

   Registro_Insere_BD = gErr

   Select Case gErr

        Case 209374, 209375
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209376)

    End Select

    Call Rotina_Erro(vbOKOnly, "Localização do Erro: " & sErro, gErr)

End Function

Private Function SpedFiscal_Grava_Tabelas(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 22) As Long, lComandoAux As Long
Dim tNFiscal As typeNFiscal, tTipoDocInfo As typeTipoDocInfo, tTributacaoDoc As typeTributacaoDoc
Dim tItemNF As typeItemNF, tTribDocItem As typeTributacaoDocItem, iIPICSTE As Integer, iIPICSTS As Integer, iICMSIncide As Integer, iISSIncide As Integer
Dim sSQLNF As String, sSQLItemNF As String
Dim objDoc As ClassSpedDocFiscais, objItem As ClassSpedDocFiscaisItens
Dim bNaoIncideICMS As Boolean, iNatureza As Integer, iCondPagto As Integer
Dim sCodVerificacaoNFe As String, sCgc As String, sChave As String
Dim tNFOrig As typeNFiscal, tTribDocOrig As typeTributacaoDoc, bExisteNFOrig As Boolean, bNFOrigPropEmp As Boolean
Dim lNumNFInut As Long, lNumNFDeInut As Long, lNumNFAteInut As Long
Dim sSerieNFInut As String, dtDataNFInut As Date, iFilialEmpNFInut As Integer, sCGCNFInut As String
Dim sNumDe As String, sNumAte As String
Dim dValorRecebido As Double, sSQLCR As String, dValorNoCR As Double
Dim iPasso As Integer, sCodAtividadeTab511 As String, dAliquotaContrPrev As Double
Dim sSQLNFBase, sSQLNFRec As String

On Error GoTo Erro_SpedFiscal_Grava_Tabelas

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
   
    sSQLNFBase = "SELECT F.CGC, N.FilialEmpresa, N.NumIntDoc, N.RPS, N.CodVerificacaoNFe, N.ChvNFe, N.NFe, N.TipoNFiscal, T.Tipo, N.Cliente, N.FilialCli, "
    sSQLNFBase = sSQLNFBase & "N.Fornecedor, N.FilialForn, N.Status, Complementar, N.Serie, N.NumNotaFiscal, "
    sSQLNFBase = sSQLNFBase & "N.DataEmissao , N.DataEntrada, N.DataSaida, N.ValorTotal, N.ClasseDocCPR, N.NumIntDocCPR, N.ValorDesconto, "
    sSQLNFBase = sSQLNFBase & "N.ValorProdutos , N.ValorFrete, N.FreteRespons, N.ValorSeguro, N.ValorOutrasDespesas, "
    sSQLNFBase = sSQLNFBase & "N.FilialEntrega , N.CodTransportadora, N.Placa, N.PlacaUF, N.VolumeQuant, "
    sSQLNFBase = sSQLNFBase & "N.PesoLiq , N.PesoBruto, T.TipoMovtoEstoque, T.NomeTelaNFiscal, "
    sSQLNFBase = sSQLNFBase & "TD.TipoTributacao , TD.IPIBase, TD.IPIValor, TD.ICMSBase, TD.ICMSValor, "
    sSQLNFBase = sSQLNFBase & "TD.ICMSSubstBase, TD.ICMSSubstValor, TD.ISSBase, TD.ISSValor, TD.PisRetido, TD.CofinsRetido, "
    sSQLNFBase = sSQLNFBase & "TD.NaturezaOPInterna, N.NumIntNotaOriginal, T.Faturamento "
    sSQLNFBase = sSQLNFBase & "FROM NFiscal AS N, TiposDocInfo AS T, TributacaoDoc AS TD, FiliaisEmpresa F "
    sSQLNFBase = sSQLNFBase & "WHERE F.FilialEmpresa = N.FilialEmpresa AND N.TipoNFiscal = T.Codigo AND TD.TipoDoc = 0 AND TD.NumIntDoc = N.NumIntDoc AND F.CGC LIKE ? "
   
    For iPasso = 1 To 2
    
        If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 1 Then
        
            'REGIME DE COMPETÊNCIA
            'NOTAS FISCAIS COM ENTRADA E SAÍDA NO PERÍODO ... BLOCOS A, C e D
            sSQLNF = sSQLNFBase & " AND "
            sSQLNF = sSQLNF & "(TD.PisCredito <> 0 OR TD.CofinsCredito <> 0 OR T.Faturamento <> 0) AND "
            sSQLNF = sSQLNF & "((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ( "
            sSQLNF = sSQLNF & "(Tipo = 1 Or Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND F.FilialEmpresa < 50 "
            sSQLNF = sSQLNF & " ORDER BY N.NumIntDoc" 'para garantir que reexecute passando pelos registros na mesma ordem...
        
        ElseIf (objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 2) Or (objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 1) Then
        
            'NOTAS FISCAIS PAGAS NO PERÍODO ... BLOCO F PARA REGIME DE CAIXA E PODE SER USADO PARA O REGIME DE COMPETÊNCIA PARA CONTRIBUIÇÃO PREVIDENCIÁRIA POR RECEITA BRUTA
            sSQLNF = sSQLNFBase & " AND "
            sSQLNF = sSQLNF & "(TD.PisCredito <> 0 OR TD.CofinsCredito <> 0 OR T.Faturamento <> 0) AND F.FilialEmpresa < 50 AND "
            
            'Pega as NFs se existir um recebimento no período
            'Se for adiantamento pega pela data de faturamento senão pela baixa (exceto baixa por perda)
            sSQLNF = sSQLNF & "N.NumIntDoc IN ( "
            sSQLNF = sSQLNF & "SELECT N.NumIntDoc "
            sSQLNF = sSQLNF & "FROM NFiscal AS N, TitulosRecTodos AS TR, ParcelasRecTodas As P, BaixasRec AS B, BaixasParcRec As BP "
            sSQLNF = sSQLNF & "WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND "
            sSQLNF = sSQLNF & "TR.NumIntDoc = P.NumIntTitulo AND BP.NumIntParcela = P.NumIntDoc AND B.NumIntBaixa = BP.NumIntBaixa AND "
            sSQLNF = sSQLNF & "B.Status <> 5 AND BP.Status <> 5 AND B.Motivo NOT IN (4,6) AND ((B.Motivo <> 5 AND B.Data BETWEEN ? AND ?) OR (B.Motivo = 5 AND N.DataEmissao BETWEEN ? AND ?))) "
        
        ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 2 Then
        
            'REGIME DE CAIXA
            'NOTAS FISCAIS GERADORAS DE RECEITA EMITIDAS NO PERÍODO .. REGISTRO 1900
            
            'Além das NFs que geraram receita no período e vão gerar os registros F500 será necessário
            'informar as NFs emitidas no próprio período para o registro 1900
            'Para o regime de competência não vai gerar o 1900 porque o arquivo é detalhado nos blocos A, C e D e
            'não consolidado no bloco F
            sSQLNF = sSQLNFBase & " AND "
            sSQLNF = sSQLNF & "T.Faturamento = 1 AND Status <> 7 AND (Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) AND (Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) AND F.FilialEmpresa < 50 "
            sSQLNF = sSQLNF & " ORDER BY N.NumIntDoc" 'para garantir que reexecute passando pelos registros na mesma ordem...

        End If
    
        If iPasso = 1 Then
            lComandoAux = alComando(2)
        Else
            lComandoAux = alComando(20)
        End If
        
        tNFiscal.sSerie = String(STRING_MAXIMO, 0)
        tNFiscal.sPlaca = String(STRING_MAXIMO, 0)
        tNFiscal.sPlacaUF = String(STRING_MAXIMO, 0)
        tTipoDocInfo.sNomeTelaNFiscal = String(STRING_MAXIMO, 0)
        tTributacaoDoc.sNaturezaOpInterna = String(STRING_MAXIMO, 0)
        tNFiscal.sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
        tNFiscal.sChvNFe = String(STRING_MAXIMO, 0)
        sCgc = String(STRING_CGC, 0)
            
        lErro = Comando_Executar(lComandoAux, sSQLNF, _
                                                sCgc, tNFiscal.iFilialEmpresa, tNFiscal.lNumIntDoc, tNFiscal.iRecibo, tNFiscal.sCodVerificacaoNFe, tNFiscal.sChvNFe, tNFiscal.iNFe, tNFiscal.iTipoNFiscal, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
                                                tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
                                                tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, tTributacaoDoc.iTipoTributacao, tTributacaoDoc.dIPIBase, tTributacaoDoc.dIPIValor, tTributacaoDoc.dICMSBase, tTributacaoDoc.dICMSValor, _
                                                tTributacaoDoc.dICMSSubstBase, tTributacaoDoc.dICMSSubstValor, tTributacaoDoc.dISSBase, tTributacaoDoc.dISSValor, tTributacaoDoc.dPISRetido, tTributacaoDoc.dCOFINSRetido, tTributacaoDoc.sNaturezaOpInterna, tNFiscal.lNumIntNotaOriginal, tTipoDocInfo.iFaturamento, _
                                                left(objEFD.objFilialEmpresa.sCgc, 8) & "%", objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 209377
                               
        lErro = Comando_BuscarPrimeiro(lComandoAux)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objDoc = New ClassSpedDocFiscais
            
            objDoc.iFilialEmpresa = tNFiscal.iFilialEmpresa
            objDoc.iTipoDoc = SPED_DOC_NF
            objDoc.lNumIntDoc = tNFiscal.lNumIntDoc
            objDoc.lNumIntSped = objEFD.lNumIntSped
            
            sChave = "NF ID: " & CStr(tNFiscal.lNumIntDoc)
            
            'Lê LivRegES
            lErro = Comando_Executar(alComando(9), "SELECT Modelo FROM LivRegES WHERE TipoNumIntDocOrigem = 0 AND NumIntDocOrigem = ? ", tTipoDocInfo.iModeloArqICMS, tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209379
        
            lErro = Comando_BuscarPrimeiro(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209380
        
            'Se não tiver registro em LivRegES não tem ICMS
            If lErro <> AD_SQL_SUCESSO Then
                tTipoDocInfo.iModeloArqICMS = 0
                bNaoIncideICMS = True
            Else
                If tTipoDocInfo.iModeloArqICMS >= 9000 Then
                    If tNFiscal.iNFe = MARCADO Then
                        tTipoDocInfo.iModeloArqICMS = 55
                    Else
                        tTipoDocInfo.iModeloArqICMS = 1
                    End If
                End If
                bNaoIncideICMS = False
            End If
            
            If right(UCase(tNFiscal.sSerie), 2) <> "-E" Then
                objDoc.sSerie = tNFiscal.sSerie
            Else
                objDoc.sSerie = left(tNFiscal.sSerie, Len(tNFiscal.sSerie) - 2)
            End If
            objDoc.lNumDocumento = tNFiscal.lNumNotaFiscal
            objDoc.dtDataEmissao = tNFiscal.dtDataEmissao
            
            If tTipoDocInfo.iTipo = 2 Then
                objDoc.sOperacao = 1 'Saída
                objDoc.dtDataES = tNFiscal.dtDataSaida
            Else
                objDoc.sOperacao = 0 'Entrada;
                objDoc.dtDataES = tNFiscal.dtDataEntrada
            End If
            
            If tTipoDocInfo.iTipo = 2 Or tTipoDocInfo.iTipo = 1 Then
                objDoc.sEmitente = "0" 'Emissão própria
            Else
                objDoc.sEmitente = "1" 'Terceiros
            End If
                   
            'Código Descrição
            '00 Documento regular
            '01 Escrituração extemporânea de documento regular
            '02 Documento cancelado
            '03 Escrituração extemporânea de documento cancelado
            '04 NF-e ou CT-e – denegado
            '05 NF-e ou CT-e - Numeração inutilizada
            '06 Documento Fiscal Complementar
            '07 Escrituração extemporânea de documento complementar
            '08 Documento Fiscal emitido com base em Regime Especial ou Norma Específica
            objDoc.iSituacao = 0
            If tNFiscal.iStatus = STATUS_CANCELADO Then
                objDoc.iSituacao = 2
                'Busca Denegado senão vai ser cancelada mesmo
                If tNFiscal.iNFe = 1 Then
        
                    'se for uma nota eletronica denegada no ambiente de producao
                    lErro = Comando_Executar(alComando(4), "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", tNFiscal.iCanal, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 209381
                
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209382
                
                    'nota fiscal denegada
                    If lErro = AD_SQL_SUCESSO Then objDoc.iSituacao = 4
                        
                End If
            Else
                If tTipoDocInfo.iComplementar = MARCADO Then
                    objDoc.iSituacao = 6
'                Else
'                    'Busca Denegado senão vai ser regular
'                    If tNFiscal.iNFe = 1 Then
'
'                        'se for uma nota eletronica denegada no ambiente de producao
'                        lErro = Comando_Executar(alComando(4), "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", tNFiscal.iCanal, tNFiscal.lNumIntDoc)
'                        If lErro <> AD_SQL_SUCESSO Then gError 209381
'
'                        lErro = Comando_BuscarPrimeiro(alComando(4))
'                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209382
'
'                        'nota fiscal denegada
'                        If lErro = AD_SQL_SUCESSO Then objDoc.iSituacao = 4
'
'                    End If
                End If
            End If
            
            Select Case tNFiscal.iFreteRespons
            
                Case 1 'Emitente
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "1"
                    Else
                        objDoc.sFrete = "0"
                    End If
                Case 2 'Destinatário
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "2"
                    Else
                        objDoc.sFrete = "1"
                    End If
                Case 3 'Terceiros
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "0"
                    Else
                        objDoc.sFrete = "2"
                    End If
                Case 4 'Sem Frete
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "9"
                    Else
                        objDoc.sFrete = "9"
                    End If
            
            End Select
            
            If tTipoDocInfo.iModeloArqICMS = 65 Then objDoc.sFrete = "9"
            
            objDoc.sModelo = Format(tTipoDocInfo.iModeloArqICMS, "00")
            
            objDoc.dValorBaseICMS = tTributacaoDoc.dICMSBase
            objDoc.dValorBaseICMSST = tTributacaoDoc.dICMSSubstBase
            objDoc.dValorBaseIPI = tTributacaoDoc.dIPIBase
            objDoc.dValorDespesas = tNFiscal.dValorOutrasDespesas
            objDoc.dValorFrete = tNFiscal.dValorFrete
            objDoc.dValorICMS = tTributacaoDoc.dICMSValor
            objDoc.dValorICMSST = tTributacaoDoc.dICMSSubstValor
            objDoc.dValorIPI = tTributacaoDoc.dIPIValor
            objDoc.dValorISS = tTributacaoDoc.dISSValor
            objDoc.dValorSeguro = tNFiscal.dValorSeguro
            objDoc.dValorTotal = tNFiscal.dValorTotal
            objDoc.dValorPisRetido = tTributacaoDoc.dPISRetido
            objDoc.dValorCofinsRetido = tTributacaoDoc.dCOFINSRetido
            objDoc.sCgc = sCgc
    
            objDoc.iFilialCli = tNFiscal.iFilialCli
            objDoc.iFilialForn = tNFiscal.iFilialForn
            objDoc.lCliente = tNFiscal.lCliente
            objDoc.lFornecedor = tNFiscal.lFornecedor
            
            'Pegar a chave
            If tNFiscal.iNFe = MARCADO Then
            
                objDoc.sChaveEletronica = tNFiscal.sChvNFe
            
                sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
            
                'Pega a chave de autorização
                lErro = Comando_Executar(alComando(7), "SELECT ChNFe FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN (?,?) AND tpAmb = 0 ORDER BY cStat", sCodVerificacaoNFe, tNFiscal.lNumIntDoc, 100, 302)
                If lErro <> AD_SQL_SUCESSO Then gError 209387
            
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209388
            
                If lErro = AD_SQL_SUCESSO Then objDoc.sChaveEletronica = sCodVerificacaoNFe
                
            ElseIf tNFiscal.iRecibo = MARCADO And Len(Trim(tNFiscal.sCodVerificacaoNFe)) = 0 Then
            
                sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
            
                'Pega a chave de autorização
                lErro = Comando_Executar(alComando(8), "SELECT CodigoVerificacao FROM RPSWebProt WHERE NumIntNF = ? AND Ambiente = 1 AND CodigoVerificacao <> ''", sCodVerificacaoNFe, tNFiscal.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 209389
            
                lErro = Comando_BuscarPrimeiro(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209390
            
                If lErro = AD_SQL_SUCESSO Then objDoc.sChaveEletronica = sCodVerificacaoNFe
                
            Else
                objDoc.sChaveEletronica = tNFiscal.sCodVerificacaoNFe 'Pode ser RPS com envio via arquivo
            End If
            
            'Pegar a condição de pagamento
            iCondPagto = 0
            Select Case tNFiscal.iClasseDocCPR
                Case CPR_TITULO_RECEBER
                
                    lErro = Comando_Executar(alComando(5), "SELECT CondicaoPagto FROM TitulosRecTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 209391
                
                    lErro = Comando_BuscarPrimeiro(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209392
                    
                Case CPR_TITULO_PAGAR
            
                    lErro = Comando_Executar(alComando(6), "SELECT CondicaoPagto FROM TitulosPagTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 209393
                
                    lErro = Comando_BuscarPrimeiro(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                    
                Case CPR_NF_PAGAR
                
                    lErro = Comando_Executar(alComando(16), "SELECT T.CondicaoPagto FROM TitulosPagTodos AS T, NFsPag_Todas As N WHERE T.NumIntDoc = N.NumIntTitPag AND N.NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 211985
                
                    lErro = Comando_BuscarPrimeiro(alComando(16))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211986
                    
            End Select
            
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '9- Sem pagamento.
            Select Case iCondPagto
                Case 0
                    objDoc.sPagamento = "9"
                Case 1
                    objDoc.sPagamento = "0"
                Case Else
                    objDoc.sPagamento = "1"
            End Select
            
            sSQLItemNF = "SELECT I.NumIntDoc, I.Item, I.Produto, I.UnidadeMed, I.Quantidade, I.PrecoUnitario, I.ValorDesconto, I.DescricaoItem,  "
            sSQLItemNF = sSQLItemNF & "T.NaturezaOP, T.TipoTributacao, T.IPITipo, T.IPIBaseCalculo, T.IPIALiquota, T.IPIValor,T.IPICredito, T.ICMSTipo, T.CST, "
            sSQLItemNF = sSQLItemNF & "T.ICMSBase, T.ICMSPercRedBase, T.ICMSAliquota, T.ICMSValor, T.ICMSCredito, T.ICMSSubstBase, T.ICMSSubstAliquota, T.ICMSSubstValor, "
            sSQLItemNF = sSQLItemNF & "T.PISCredito, T.CofinsCredito, T.ICMSSubstPercRedBase, T.ICMSSubstPercMVA, T.PISTipo, T.PISBase, T.PISAliquota, T.PISAliquotaValor, T.PISQtde, T.PISValor, "
            sSQLItemNF = sSQLItemNF & "T.PISSTBase, T.PISSTAliquota, T.PISSTAliquotaValor, T.PISSTQtde, T.PISSTValor, T.CofinsTipo, T.CofinsBase, T.CofinsAliquota, T.CofinsAliquotaValor, T.CofinsQtde, T.CofinsValor, "
            sSQLItemNF = sSQLItemNF & "T.CofinsSTBase, T.CofinsSTAliquota, T.CofinsSTAliquotaValor, T.CofinsSTQtde, T.CofinsSTValor, T.ISSBase, T.ISSAliquota, T.ISSValor, T.IPIEnquadramentoCodigo, "
            sSQLItemNF = sSQLItemNF & "T.ValorFreteItem , T.ValorSeguroItem, T.ValorOutrasDespesasItem, T.OrigemMercadoria, TT.CSTEntrada, TT.CSTSaida, TM.ICMSIncide, TM.ISSIncide, P.Natureza, T.NatBCCred, T.IPICodProduto "
            sSQLItemNF = sSQLItemNF & "FROM ItensNFiscal AS I, TributacaoDocItem AS T, TiposTribIPI As TT, TiposDeTributacaoMovto As TM, Produtos AS P "
            sSQLItemNF = sSQLItemNF & "WHERE P.Codigo = I.Produto AND TM.Tipo = T.TipoTributacao AND TT.Tipo = T.IPITipo AND T.Tipodoc = 0 AND T.NumIntDocItem = I.NumIntDoc AND I.NumIntNF = ? ORDER BY I.Item "
            
            tItemNF.sProduto = String(STRING_MAXIMO, 0)
            tItemNF.sUnidadeMed = String(STRING_MAXIMO, 0)
            tItemNF.sDescricaoItem = String(STRING_MAXIMO, 0)
            tTribDocItem.sNaturezaOp = String(STRING_MAXIMO, 0)
            tTribDocItem.sCST = String(STRING_MAXIMO, 0)
            tTribDocItem.sIPIEnquadramentoCodigo = String(STRING_MAXIMO, 0)
            tTribDocItem.sNatBCCred = String(STRING_MAXIMO, 0)
            tTribDocItem.sIPICodProduto = String(STRING_MAXIMO, 0)
            
            lErro = Comando_Executar(alComando(3), sSQLItemNF, _
                                                    tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tItemNF.sDescricaoItem, _
                                                    tTribDocItem.sNaturezaOp, tTribDocItem.iTipoTributacao, tTribDocItem.iIPITipo, tTribDocItem.dIPIBaseCalculo, tTribDocItem.dIPIAliquota, tTribDocItem.dIPIValor, tTribDocItem.dIPICredito, tTribDocItem.iICMSTipo, tTribDocItem.sCST, _
                                                    tTribDocItem.dICMSBase, tTribDocItem.dICMSPercRedBase, tTribDocItem.dICMSAliquota, tTribDocItem.dICMSValor, tTribDocItem.dICMSCredito, tTribDocItem.dICMSSubstBase, tTribDocItem.dICMSSubstAliquota, tTribDocItem.dICMSSubstValor, _
                                                    tTribDocItem.dPISCredito, tTribDocItem.dCOFINSCredito, tTribDocItem.dICMSSubstPercRedBase, tTribDocItem.dICMSSubstPercMVA, tTribDocItem.iPISTipo, tTribDocItem.dPISBase, tTribDocItem.dPISAliquota, tTribDocItem.dPISAliquotaValor, tTribDocItem.dPISQtde, tTribDocItem.dPISValor, _
                                                    tTribDocItem.dPISSTBase, tTribDocItem.dPISSTAliquota, tTribDocItem.dPISSTAliquotaValor, tTribDocItem.dPISSTQtde, tTribDocItem.dPISSTValor, tTribDocItem.iCOFINSTipo, tTribDocItem.dCOFINSBase, tTribDocItem.dCOFINSAliquota, tTribDocItem.dCOFINSAliquotaValor, tTribDocItem.dCOFINSQtde, tTribDocItem.dCOFINSValor, _
                                                    tTribDocItem.dCOFINSSTBase, tTribDocItem.dCOFINSSTAliquota, tTribDocItem.dCOFINSSTAliquotaValor, tTribDocItem.dCOFINSSTQtde, tTribDocItem.dCOFINSSTValor, tTribDocItem.dISSBase, tTribDocItem.dISSAliquota, tTribDocItem.dISSValor, tTribDocItem.sIPIEnquadramentoCodigo, _
                                                    tTribDocItem.dValorFreteItem, tTribDocItem.dValorSeguroItem, tTribDocItem.dValorOutrasDespesasItem, tTribDocItem.iOrigemMercadoria, iIPICSTE, iIPICSTS, iICMSIncide, iISSIncide, iNatureza, tTribDocItem.sNatBCCred, tTribDocItem.sIPICodProduto, _
                                                    tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209395
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209396
        
            Do While lErro <> AD_SQL_SEM_DADOS
            
                sChave = "NF ID: " & CStr(tNFiscal.lNumIntDoc) & " Item: " & CStr(tItemNF.iItem)
            
                Set objItem = New ClassSpedDocFiscaisItens
                objDoc.colItens.Add objItem
                
                objItem.dPrecoUnitario = tItemNF.dPrecoUnitario
                objItem.dCOFINSAliquota = tTribDocItem.dCOFINSAliquota
                objItem.dCofinsAliquotaQtd = tTribDocItem.dCOFINSAliquotaValor
                objItem.dCofinsBCQtd = tTribDocItem.dCOFINSQtde
                objItem.dICMSAliquota = tTribDocItem.dICMSAliquota
                objItem.dICMSSTAliquota = tTribDocItem.dICMSSubstAliquota
                objItem.dIPIAliquota = tTribDocItem.dIPIAliquota
                objItem.dPISAliquota = tTribDocItem.dPISAliquota
                objItem.dPISAliquotaQtd = tTribDocItem.dPISAliquotaValor
                objItem.dPISBCQtd = tTribDocItem.dPISQtde
                objItem.dQuantidade = tItemNF.dQuantidade
                objItem.dValorBaseCofins = tTribDocItem.dCOFINSBase
                objItem.dValorBaseICMS = tTribDocItem.dICMSBase
                objItem.dValorBaseICMSST = tTribDocItem.dICMSSubstBase
                objItem.dValorBaseIPI = tTribDocItem.dIPIBaseCalculo
                objItem.dValorBasePIS = tTribDocItem.dPISBase
                objItem.dValorCofins = tTribDocItem.dCOFINSValor
                'objItem.dValorCofinsRetido
                objItem.dValorCofinsST = tTribDocItem.dCOFINSSTValor
                objItem.dValorDesconto = tItemNF.dValorDesconto
                objItem.dValorDespesas = tTribDocItem.dValorOutrasDespesasItem
                objItem.dValorFrete = tTribDocItem.dValorFreteItem
                objItem.dValorICMS = tTribDocItem.dICMSValor
                objItem.dValorICMSST = tTribDocItem.dICMSSubstValor
                objItem.dValorIPI = tTribDocItem.dIPIValor
                objItem.dValorISS = tTribDocItem.dISSValor
                'objItem.dValorNaoTributado
                objItem.dValorPIS = tTribDocItem.dPISValor
                'objItem.dValorPisRetido
                objItem.dValorPisST = tTribDocItem.dPISSTValor
                objItem.dValorSeguro = tTribDocItem.dValorSeguroItem
                objItem.iItem = tItemNF.iItem
                objItem.sCFOP = tTribDocItem.sNaturezaOp
                objItem.sNatBCCred = tTribDocItem.sNatBCCred
                objItem.sNCM = tTribDocItem.sIPICodProduto
                
                '"A majoração da alíquota em referência, pela Medida provisória nº 563/2012, foi tão somente no cálculo da Cofins na importação, especificada no art. 8º da Lei nº 10.865/04.
                'Referido aumento não se estendeu aos créditos na importação, que continuam tendo o seu valor determinado de acordo com o disposto no art. 15 da referida Lei, o qual estabelece em seu § 3º que para o cálculo dos créditos na importação aplicam-se as alíquotas dispostas no caput do art. 2º das Leis nº 10.637/02 (PIS) e nº 10.833/03 (Cofins).
                'O objetivo do aumento em 1% foi o de desestimular a importação dos produtos listados e sujeitos à Contribuição Previdenciária sobre a Receita Bruta (Lei nº 12.546). Caso o aumento fosse repassado para os créditos, teria os seus efeitos anulados.
                'Desta forma:
                'Considerando que na EFD Contribuições não se escritura as operações referentes ao PIS/Pasep - Importação e à Cofins - Importação,  e sim e tão somente, as operações referentes aos créditos das referidas importações;
                'Considerando que na apuração dos créditos sobre a importação, as alíquotas aplicáveis (1,65% e 7,6%) não sofreram alterações.
                'Resta esclarecido que o aumento em 1% na alíquota da Cofins-Importação, não é passível de escrituração da EFD-Contribuições.
                'OBS; Os créditos nas importações, previstos no art. 15 da Lei nº 10.865/04, não correspondem exatamente aos valores pagos na importação e sim, corresponde aos valores determinados conforme as disposições do referido artigo - Base de Cálculo x Alíquota = Valor do Crédito."
                'se é uma importação com Cofins de 8,6% ajusta pois o 1% a mais não constitui crédito
                If left(objItem.sCFOP, 1) = "3" And Abs(objItem.dCOFINSAliquota - 0.086) < QTDE_ESTOQUE_DELTA Then
                    objItem.dCOFINSAliquota = 0.076
                    objItem.dValorCofins = Arredonda_Moeda(objItem.dValorBaseCofins * objItem.dCOFINSAliquota)
                End If
                
                'Se paga contribuição previdenciária pela receita bruta
                If objEFD.objFilialEmpresa.iContribPrevPorRecBruta = MARCADO Then
                    objItem.dAliquotaContrPrev = objEFD.objFilialEmpresa.dAliquotaPadraoContribPrev
                    If Len(Trim(objItem.sNCM)) > 0 Then
                    
                        sCodAtividadeTab511 = String(STRING_MAXIMO, 0)
                    
                        lErro = Comando_Executar(alComando(21), "SELECT CodAtividadeTab511, AliquotaContrPrev FROM CodAtivContrPrev WHERE DataIniEscrituracao <= ? AND DataFimEscrituracao >= ? AND NCM = LEFT(?,LEN(NCM)) AND ExNCM NOT LIKE ?", sCodAtividadeTab511, dAliquotaContrPrev, objDoc.dtDataEmissao, objDoc.dtDataEmissao, objItem.sNCM, "%" & objItem.sNCM & "%")
                        If lErro <> AD_SQL_SUCESSO Then gError 211987
                        
                        lErro = Comando_BuscarPrimeiro(alComando(21))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211988
                    
                        If lErro <> AD_SQL_SEM_DADOS Then
                            objItem.sCodAtividadeTab511 = sCodAtividadeTab511
                            objItem.dAliquotaContrPrev = dAliquotaContrPrev
                        Else
                            objItem.sCodAtividadeTab511 = "99999999" 'Código Genérico - Contribuição Previdenciária sobre Receitas
                        End If
                    
                    Else
                        If objDoc.sModelo = "00" Or iISSIncide = MARCADO Then
                            objItem.sCodAtividadeTab511 = "00009999" 'Código Genérico - Serviços
                        End If
                    End If
                End If
                
                If tTipoDocInfo.iTipo = 2 Or tTribDocItem.iCOFINSTipo >= 50 Then
                    objItem.sCofinsCST = Format(tTribDocItem.iCOFINSTipo, "00")
                Else
                    Select Case tTribDocItem.iCOFINSTipo
                        Case 1
                            objItem.sCofinsCST = "56"
                        Case 6
                            objItem.sCofinsCST = "73"
                        Case 8
                            objItem.sCofinsCST = "74"
                        Case 9
                            objItem.sCofinsCST = "72"
                        Case 49
                            objItem.sCofinsCST = "98"
                        Case Else
                            objItem.sCofinsCST = Format(tTribDocItem.iCOFINSTipo, "00")
                    End Select
                End If
                
                objItem.sDescricao = tItemNF.sDescricaoItem
                objItem.sICMSCST = tTribDocItem.sCST
                objItem.sIPICodEnq = Format(tTribDocItem.sIPIEnquadramentoCodigo, "000")
                
                If tTipoDocInfo.iTipo = 2 Then
                    objItem.sIPICST = Format(iIPICSTS, "00") 'Saída
                Else
                    objItem.sIPICST = Format(iIPICSTE, "00") 'Entrada
                End If
                            
                If tTipoDocInfo.iTipo = 2 Or tTribDocItem.iPISTipo >= 50 Then
                    objItem.sPISCST = Format(tTribDocItem.iPISTipo, "00")
                Else
                    Select Case tTribDocItem.iPISTipo
                        Case 1
                            objItem.sPISCST = "56"
                        Case 6
                            objItem.sPISCST = "73"
                        Case 8
                            objItem.sPISCST = "74"
                        Case 9
                            objItem.sPISCST = "72"
                        Case 49
                            objItem.sPISCST = "98"
                        Case Else
                            objItem.sPISCST = Format(tTribDocItem.iPISTipo, "00")
                    End Select
                End If
                
                objItem.sProduto = tItemNF.sProduto
                objItem.sUM = tItemNF.sUnidadeMed
                
                'Movimentação física do ITEM/PRODUTO:
                '0. SIM
                '1. NÃO
                If tTipoDocInfo.iTipoMovtoEstoque = 0 Then
                    objItem.sMovFisica = "1"
                Else
                    objItem.sMovFisica = "0"
                End If
                
                objItem.dPrecoTotal = objItem.dQuantidade * objItem.dPrecoUnitario
                If iNatureza = NATUREZA_PROD_SERVICO Then
                    objItem.dValorServico = objItem.dPrecoTotal
                    objItem.sMovFisica = "1"
                Else
                    objDoc.dValorMercadoria = objDoc.dValorMercadoria + objItem.dPrecoTotal
                End If
                
                objDoc.dValorBaseCofins = objDoc.dValorBaseCofins + objItem.dValorBaseCofins
                objDoc.dValorBasePIS = objDoc.dValorBasePIS + objItem.dValorBasePIS
                objDoc.dValorCofins = objDoc.dValorCofins + objItem.dValorCofins
                objDoc.dValorCofinsRetido = objDoc.dValorCofinsRetido + objItem.dValorCofinsRetido
                objDoc.dValorCofinsST = objDoc.dValorCofinsST + objItem.dValorCofinsST
                objDoc.dValorDesconto = objDoc.dValorDesconto + objItem.dValorDesconto
                objDoc.dValorPIS = objDoc.dValorPIS + objItem.dValorPIS
                objDoc.dValorPisRetido = objDoc.dValorPisRetido + objItem.dValorPisRetido
                objDoc.dValorPisST = objDoc.dValorPisST + objItem.dValorPisST
                objDoc.dValorServico = objDoc.dValorServico + objItem.dValorServico
                objDoc.dValorNaoTributado = objDoc.dValorNaoTributado + objItem.dValorNaoTributado
            
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209397
                        
            Loop
            
            'Tabela de modelo de documentos fiscais e registros correspondentes na EFD-PIS/Cofins:
            'Página 7/8 do manual
            'Guia Prático EFD-PIS/COFINS – Versão 1.00
            'Atualização: 31 de janeiro de 2011
            If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 1 Then
                If objDoc.sModelo = "00" Or objDoc.sModelo = "01" Or objDoc.sModelo = "1B" Or objDoc.sModelo = "04" Or objDoc.sModelo = "55" Or objDoc.sModelo = "65" Then
                    If tTipoDocInfo.iFaturamento = 2 And objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = DESMARCADO Then
                        'Se é uma devolução no regime cumulativo
                        objDoc.sBloco = "M"
                        objDoc.sRegistro = "M220"
                    Else
                        If iISSIncide = MARCADO Or objDoc.sModelo = "00" Then
                            objDoc.sBloco = "A"
                            objDoc.sRegistro = "A100"
                        Else
                            objDoc.sBloco = "C"
                            objDoc.sRegistro = "C100"
                        End If
                    End If
                ElseIf objDoc.sModelo = "02" Then
                    objDoc.sBloco = "C"
                    objDoc.sRegistro = "C380"
                ElseIf objDoc.sModelo = "06" Or objDoc.sModelo = "28" Or objDoc.sModelo = "29" Then
                    objDoc.sBloco = "C"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "C600" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "C500" 'Aquisição:
                    End If
                ElseIf objDoc.sModelo = "07" Or objDoc.sModelo = "08" Or objDoc.sModelo = "09" Or objDoc.sModelo = "10" Or objDoc.sModelo = "11" Or objDoc.sModelo = "26" Or objDoc.sModelo = "27" Or objDoc.sModelo = "57" Or objDoc.sModelo = "8B" Then
                    objDoc.sBloco = "D"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "D200" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "D100" 'Aquisição:
                    End If
                ElseIf objDoc.sModelo = "21" Or objDoc.sModelo = "22" Then
                    objDoc.sBloco = "D"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "D600" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "D500" 'Aquisição:
                    End If
                End If
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 2 Then
                objDoc.sBloco = "" 'Deixa vazio pois vai servir só para apurar a Receita, mas não vai gerar registro no bloco F
                objDoc.sRegistro = "F500" 'COMPOSIÇÃO DA RECEITA ESCRITURADA NO PERÍODO
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 1 Then
                objDoc.sBloco = "F"
                objDoc.sRegistro = "F500" 'COMPOSIÇÃO DA RECEITA ESCRITURADA NO PERÍODO
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 2 Then
                objDoc.sBloco = "1"
                objDoc.sRegistro = "1900" 'CONSOLIDAÇÃO DOS DOCUMENTOS EMITIDOS NO PERÍODO GERADORES DE RECEITA
                If objDoc.sModelo = "00" Or iISSIncide = MARCADO Then
                    objDoc.sModelo = "98" 'Nota Fiscal de Prestação de Serviços (ISSQN)
                End If
            End If
            
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '9- Sem pagamento
            'C100
            'Obs.: A partir de 01/07/2012 passará a ser:
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '2 - Outros
            If objEFD.iCodVersao >= 3 Then ' versão posterior a 01/07/2012
                If objDoc.sPagamento = "9" And objDoc.sBloco <> "A" Then
                    objDoc.sPagamento = "2"
                End If
            End If
            
            If tNFiscal.iClasseDocCPR = CPR_TITULO_RECEBER Then
            
                'Acumula o total recebido por cada nf (depois vai ratear por CST/Alíquota)
                sSQLCR = "SELECT SUM(BP.ValorRecebido) "
                sSQLCR = sSQLCR & "FROM NFiscal AS N, TitulosRecTodos AS TR, ParcelasRecTodas As P, BaixasRec AS B, BaixasParcRec As BP "
                sSQLCR = sSQLCR & "WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND "
                sSQLCR = sSQLCR & "TR.NumIntDoc = P.NumIntTitulo AND BP.NumIntParcela = P.NumIntDoc AND B.NumIntBaixa = BP.NumIntBaixa AND "
                sSQLCR = sSQLCR & "B.Status <> 5 AND BP.Status <> 5 AND B.Motivo NOT IN (4,6) AND ((B.Motivo <> 5 AND B.Data BETWEEN ? AND ?) OR (B.Motivo = 5 AND N.DataEmissao BETWEEN ? AND ?)) AND N.NumIntDoc = ? "
    
                lErro = Comando_Executar(alComando(18), sSQLCR, _
                                                        dValorRecebido, objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim, objDoc.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211989
                
                lErro = Comando_BuscarPrimeiro(alComando(18))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211990
    
                objDoc.dValorRecebido = dValorRecebido
                
                lErro = Comando_Executar(alComando(19), "SELECT TR.Valor FROM NFiscal AS N, TitulosRecTodos AS TR WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND N.NumIntDoc = ?", dValorNoCR, objDoc.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211991
                
                lErro = Comando_BuscarPrimeiro(alComando(19))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211992
                
                If lErro = AD_SQL_SUCESSO Then objDoc.dValorNoCR = dValorNoCR
                
            End If
            
            If objDoc.sRegistro = "D100" Then
            
                'Indicador da Natureza do Frete Contratado, referente a:
                '0 - Operações de vendas, com ônus suportado pelo estabelecimento vendedor;
                '1 - Operações de vendas, com ônus suportado pelo adquirente;
                '2 - Operações de compras (bens para revenda, matérias-prima e outros produtos, geradores de crédito);
                '3 - Operações de compras (bens para revenda, matérias-prima e outros produtos, não geradores de crédito);
                '4 - Transferência de produtos acabados entre estabelecimentos da pessoa jurídica;
                '5 - Transferência de produtos em elaboração entre estabelecimentos da pessoa jurídica
                '9 - Outras.
                
                bExisteNFOrig = False
                bNFOrigPropEmp = False
                
                If tNFiscal.lNumIntNotaOriginal <> 0 Then
                
                    tNFOrig.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                
                    lErro = Comando_Executar(alComando(13), "SELECT N.Cliente, N.FilialCli, N.Fornecedor, N.FilialForn, N.ValorFrete, N.NaturezaOP, T.PisCredito, T.CofinsCredito FROM NFiscal AS N, TributacaoDoc AS T WHERE T.TipoDoc = 0 AND T.NumIntDoc = N.NumIntDoc AND N.NumIntDoc = ?", tNFOrig.lCliente, tNFOrig.iFilialCli, tNFOrig.lFornecedor, tNFOrig.iFilialForn, tNFOrig.dValorFrete, tNFOrig.sNaturezaOp, tTribDocOrig.dPISCredito, tTribDocOrig.dCOFINSCredito, tNFiscal.lNumIntNotaOriginal)
                    If lErro <> AD_SQL_SUCESSO Then gError 209393
                
                    lErro = Comando_BuscarPrimeiro(alComando(13))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                    
                    If lErro = AD_SQL_SUCESSO Then
                        
                        bExisteNFOrig = True
                        
                        If tNFOrig.lCliente <> 0 Then
                        
                            sCgc = String(STRING_CGC, 0)
                        
                            lErro = Comando_Executar(alComando(14), "SELECT CGC FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", sCgc, tNFOrig.lCliente, tNFOrig.iFilialCli)
                            If lErro <> AD_SQL_SUCESSO Then gError 209393
                        
                            lErro = Comando_BuscarPrimeiro(alComando(14))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                            
                            If left(sCgc, 8) = left(objEFD.objFilialEmpresa.sCgc, 8) Then bNFOrigPropEmp = True
                        
                        Else
                            
                            sCgc = String(STRING_CGC, 0)
                        
                            lErro = Comando_Executar(alComando(15), "SELECT CGC FROM FiliaisFornecedores WHERE CodFornecedor = ? AND CodFilial = ?", sCgc, tNFOrig.lFornecedor, tNFOrig.iFilialForn)
                            If lErro <> AD_SQL_SUCESSO Then gError 209393
                        
                            lErro = Comando_BuscarPrimeiro(alComando(15))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                            
                            If left(sCgc, 8) = left(objEFD.objFilialEmpresa.sCgc, 8) Then bNFOrigPropEmp = True
                        
                        End If
                        
                    End If
                
                End If
    
                If Not bExisteNFOrig Then
                    '9 - Outras.
                    objDoc.sIndNatFrtPis = "9"
                    objDoc.sIndNatFrtCofins = "9"
                Else
                    If bNFOrigPropEmp Then 'Se a NF original é da própria empresa então é uma transferência
                        Select Case tNFOrig.sNaturezaOp
                            Case ""
                                '5 - Transferência de produtos em elaboração entre estabelecimentos da pessoa jurídica
                                objDoc.sIndNatFrtPis = "5"
                                objDoc.sIndNatFrtCofins = "5"
                            Case Else
                                '4 - Transferência de produtos acabados entre estabelecimentos da pessoa jurídica;
                                objDoc.sIndNatFrtPis = "4"
                                objDoc.sIndNatFrtCofins = "4"
                        End Select
                    Else
                        If tNFiscal.iTipoNFiscal = DOCINFO_CFECT Then 'Compra
                            '2 - Operações de compras (bens para revenda, matérias-prima e outros produtos, geradores de crédito);
                            '3 - Operações de compras (bens para revenda, matérias-prima e outros produtos, não geradores de crédito);
                            If tTribDocOrig.dPISCredito > 0 Then
                                objDoc.sIndNatFrtPis = "2"
                            Else
                                objDoc.sIndNatFrtPis = "3"
                            End If
                            If tTribDocOrig.dCOFINSCredito > 0 Then
                                objDoc.sIndNatFrtCofins = "2"
                            Else
                                objDoc.sIndNatFrtCofins = "3"
                            End If
                            
                        ElseIf tNFiscal.iTipoNFiscal = DOCINFO_CFEV Then 'Venda
                            '0 - Operações de vendas, com ônus suportado pelo estabelecimento vendedor;
                            '1 - Operações de vendas, com ônus suportado pelo adquirente;
                            If tNFOrig.dValorFrete > 0 Then
                                objDoc.sIndNatFrtPis = "1"
                                objDoc.sIndNatFrtCofins = "1"
                            Else
                                objDoc.sIndNatFrtPis = "0"
                                objDoc.sIndNatFrtCofins = "0"
                            End If
                        Else
                            '9 - Outras.
                            objDoc.sIndNatFrtPis = "9"
                            objDoc.sIndNatFrtCofins = "9"
                        End If
                    End If
                
                End If
    
            End If
            
            For Each objItem In objDoc.colItens
            
                'Preenche a Natureza caso seja um CST com direito a crédito
                If Len(Trim(objItem.sNatBCCred)) = 0 And ((objItem.sCofinsCST >= "50" And objItem.sCofinsCST <= "56") Or (objItem.sPISCST >= "50" And objItem.sPISCST <= "56")) Then
                
                    'Código Descrição
                    '01  Aquisição de bens para revenda
                    '02  Aquisição de bens utilizados como insumo
                    '03  Aquisição de serviços utilizados como insumo
                    '04  Energia elétrica e térmica, inclusive sob a forma de vapor
                    '05  Aluguéis de prédios
                    '06  Aluguéis de máquinas e equipamentos
                    '07  Armazenagem de mercadoria e frete na operação de venda
                    '08  Contraprestações de arrendamento mercantil
                    '09  Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito sobre encargos de depreciação).
                    '10  Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito com base no valor de aquisição).
                    '11  Amortização e Depreciação de edificações e benfeitorias em imóveis
                    '12  Devolução de Vendas Sujeitas à Incidência Não-Cumulativa
                    '13  Outras Operações com Direito a Crédito
                    '14  Atividade de Transporte de Cargas - Subcontratação
                    '15  Atividade Imobiliária - Custo Incorrido de Unidade Imobiliária
                    '16  Atividade Imobiliária - Custo Orçado de unidade não concluída
                    '17  Atividade de Prestação de Serviços de Limpeza, Conservação e Manutenção - vale-transporte, vale-refeição ou vale-alimentação, fardamento ou uniforme.
                    '18  Estoque de abertura de bens
                
                    Select Case objDoc.sRegistro
                
                        Case "A100"
                            objItem.sNatBCCred = "03"
                
                        Case "C500"
                            objItem.sNatBCCred = "04"
                
                        Case "D100"
                            objItem.sNatBCCred = "14"
                
                        Case "D500"
                            objItem.sNatBCCred = "03"
                
                    End Select
                
                End If
            
            Next
            
            If (objDoc.sRegistro = "F500" Or objDoc.sRegistro = "1900") And objDoc.dValorNoCR <> 0 Then
                lErro = SpedFiscal_Trata_ValorRecebido(objDoc, objEFD)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            lErro = SpedFiscal_Docs_Grava(alComando(0), alComando(1), alComando(22), objDoc, objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = Comando_BuscarProximo(lComandoAux)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209398
        
        Loop
    Next
    
    'NFS INUTILIZADAS
    sSerieNFInut = String(STRING_SERIE, 0)
    sCGCNFInut = String(STRING_CGC, 0)
    sNumDe = String(STRING_MAXIMO, 0)
    sNumAte = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(17), "SELECT F.CGC, F.FilialEmpresa, N.Data, N.Serie, N.nNFIni, N.nNFFim  FROM NFeFedRetInutNFe AS N, FiliaisEmpresa AS F WHERE N.FilialEMpresa = F.FilialEMpresa AND N.Data BETWEEN ? AND ? AND N.tpAmb = ? AND N.cStat = ? AND N.FilialEmpresa < 50 ORDER BY Data, nNFIni ", sCGCNFInut, iFilialEmpNFInut, dtDataNFInut, sSerieNFInut, sNumDe, sNumAte, objEFD.dtDataIni, objEFD.dtDataFim, 0, "102")
    If lErro <> AD_SQL_SUCESSO Then gError 209377
    
    lErro = Comando_BuscarPrimeiro(alComando(17))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378
    
    Do While lErro <> AD_SQL_SEM_DADOS

        lNumNFDeInut = StrParaLong(sNumDe)
        lNumNFAteInut = StrParaLong(sNumAte)

        For lNumNFInut = lNumNFDeInut To lNumNFAteInut
    
            Set objDoc = New ClassSpedDocFiscais
            
            objDoc.iFilialEmpresa = tNFiscal.iFilialEmpresa
            objDoc.iTipoDoc = SPED_DOC_NF_INUT
            objDoc.lNumIntDoc = lNumNFInut
            objDoc.lNumIntSped = objEFD.lNumIntSped
            objDoc.lNumDocumento = lNumNFInut
            If right(UCase(sSerieNFInut), 2) <> "-E" Then
                objDoc.sSerie = sSerieNFInut
            Else
                objDoc.sSerie = left(sSerieNFInut, Len(sSerieNFInut) - 2)
            End If
            objDoc.sModelo = "55"
            objDoc.sOperacao = 1 'Saída
            objDoc.sEmitente = "0" 'Emissão própria
            objDoc.iSituacao = 5 '05 NF-e ou CT-e - Numeração inutilizada
            objDoc.iFilialEmpresa = iFilialEmpNFInut
            objDoc.sCgc = sCGCNFInut
            objDoc.sBloco = "C"
            objDoc.sRegistro = "C100"
        
            lErro = SpedFiscal_Docs_Grava(alComando(0), alComando(1), alComando(22), objDoc, objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
        
        lErro = Comando_BuscarProximo(alComando(17))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378

    Loop
    'CUPOM FISCAL
    
    'CONHECIMENTO DE TRANSPORTE
    
    'OUTROS (Bloco F)

    SpedFiscal_Grava_Tabelas = SUCESSO

    Exit Function

Erro_SpedFiscal_Grava_Tabelas:

    SpedFiscal_Grava_Tabelas = gErr

    Select Case gErr
   
        Case ERRO_SEM_MENSAGEM
        
        Case 209377, 209378, 209398, 211987 To 211992
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 209379, 209380
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)

        Case 209381, 209382, 209387, 209388
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)

        Case 209383, 209384
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, tNFiscal.lCliente, tNFiscal.iFilialCli)

        Case 209385, 209386
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)

        Case 209389, 209390
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBPROT", gErr)

        Case 209391, 209392
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTODOS", gErr)

        Case 209393, 209394, 211985, 211986
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGTODOS", gErr)

        Case 209395 To 209397
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209399)

    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "SpedFiscal_Grava_Tabelas", sChave)
    
    Exit Function

End Function

Private Function SpedFiscal_Docs_Grava(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objDoc As ClassSpedDocFiscais, ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long, lAux As Long
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_SpedFiscal_Docs_Grava

    lErro = Comando_Executar(lComando1, "INSERT INTO SpedDocFiscais( NumIntSped, TipoDoc, NumIntDoc, Bloco, " & _
                    "Registro, FilialEmpresa, Cliente, FilialCli, Fornecedor, " & _
                    "FilialForn, DataEmissao, DataES, ValorTotal, Operacao, " & _
                    "Emitente, Modelo, Situacao, Serie, NumDocumento, " & _
                    "ChaveEletronica, Frete, Pagamento, ValorDesconto, ValorFrete, " & _
                    "ValorSeguro, ValorDespesas, ValorServico, ValorNaoTributado, ValorBaseICMS, " & _
                    "ValorICMS, ValorBaseICMSST, ValorICMSST, ValorBaseIPI, ValorIPI, " & _
                    "ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, ValorPisRetido, " & _
                    "ValorPisST, ValorCofinsRetido, ValorCofinsST, ValorISS, CGC, CodMensagem, ValorMercadoria, IndNatFrtPis, IndNatFrtCofins, ValorRecebido, ValorNoCR, ValorRecContrPrev) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc, objDoc.sBloco, _
                    objDoc.sRegistro, objDoc.iFilialEmpresa, objDoc.lCliente, objDoc.iFilialCli, objDoc.lFornecedor, _
                    objDoc.iFilialForn, objDoc.dtDataEmissao, objDoc.dtDataES, objDoc.dValorTotal, objDoc.sOperacao, _
                    objDoc.sEmitente, objDoc.sModelo, objDoc.iSituacao, objDoc.sSerie, objDoc.lNumDocumento, _
                    objDoc.sChaveEletronica, objDoc.sFrete, objDoc.sPagamento, objDoc.dValorDesconto, objDoc.dValorFrete, _
                    objDoc.dValorSeguro, objDoc.dValorDespesas, objDoc.dValorServico, objDoc.dValorNaoTributado, objDoc.dValorBaseICMS, _
                    objDoc.dValorICMS, objDoc.dValorBaseICMSST, objDoc.dValorICMSST, objDoc.dValorBaseIPI, objDoc.dValorIPI, _
                    objDoc.dValorBasePIS, objDoc.dValorPIS, objDoc.dValorBaseCofins, objDoc.dValorCofins, objDoc.dValorPisRetido, _
                    objDoc.dValorPisST, objDoc.dValorCofinsRetido, objDoc.dValorCofinsST, objDoc.dValorISS, objDoc.sCgc, objDoc.lCodMensagem, objDoc.dValorMercadoria, objDoc.sIndNatFrtPis, objDoc.sIndNatFrtCofins, objDoc.dValorRecebido, objDoc.dValorNoCR, objDoc.dValorRecContrPrev)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209400
    
    'Verifica se os itens já foram inseridos
    lErro = Comando_Executar(lComando3, "SELECT NumIntSped FROM SpedDocFiscaisItens WHERE NumIntSped = ? AND TipoDoc = ? AND NumIntDoc = ?", _
    lAux, objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc)
    If lErro <> SUCESSO Then gError 209400

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209400
    
    'Para não causar repetições de NFs de registros de emissão e recebimento ocorridos no mesmo período
    If lErro <> AD_SQL_SUCESSO Then
        
        For Each objItem In objDoc.colItens
            
            objItem.lNumIntSped = objDoc.lNumIntSped
            objItem.lNumIntDoc = objDoc.lNumIntDoc
            objItem.iTipoDoc = objDoc.iTipoDoc
        
            lErro = Comando_Executar(lComando2, "INSERT INTO SpedDocFiscaisItens( NumIntSped, TipoDoc, NumIntDoc, Item, " & _
                    "Registro, Produto, Descricao, Quantidade, UM, " & _
                    "MovFisica, ICMSCST, CFOP, ICMSAliquota, ICMSSTAliquota, " & _
                    "IPICST, IPICodEnq, IPIALiquota, PISCST, PISAliquota, " & _
                    "PISBCQtd, PISAliquotaQtd, CofinsCST, CofinsAliquota, CofinsBCQtd, " & _
                    "CofinsAliquotaQtd, ValorDesconto, ValorFrete, ValorSeguro, ValorDespesas, " & _
                    "ValorServico, ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, " & _
                    "ValorICMSST, ValorBaseIPI, ValorIPI, ValorBasePIS, ValorPIS, " & _
                    "ValorBaseCofins, ValorCofins, ValorPisRetido, ValorPisST, ValorCofinsRetido, " & _
                    "ValorCofinsST, ValorISS, PrecoUnitario, PrecoTotal, NatBCCred, ValorRecebido, ValorNoCR, NCM, CodAtividadeTab511, AliquotaContrPrev, ValorRecContrPrev) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objItem.lNumIntSped, objItem.iTipoDoc, objItem.lNumIntDoc, objItem.iItem, _
                    objItem.sRegistro, objItem.sProduto, objItem.sDescricao, objItem.dQuantidade, objItem.sUM, _
                    objItem.sMovFisica, objItem.sICMSCST, objItem.sCFOP, objItem.dICMSAliquota, objItem.dICMSSTAliquota, _
                    objItem.sIPICST, objItem.sIPICodEnq, objItem.dIPIAliquota, objItem.sPISCST, objItem.dPISAliquota, _
                    objItem.dPISBCQtd, objItem.dPISAliquotaQtd, objItem.sCofinsCST, objItem.dCOFINSAliquota, objItem.dCofinsBCQtd, _
                    objItem.dCofinsAliquotaQtd, objItem.dValorDesconto, objItem.dValorFrete, objItem.dValorSeguro, objItem.dValorDespesas, _
                    objItem.dValorServico, objItem.dValorNaoTributado, objItem.dValorBaseICMS, objItem.dValorICMS, objItem.dValorBaseICMSST, _
                    objItem.dValorICMSST, objItem.dValorBaseIPI, objItem.dValorIPI, objItem.dValorBasePIS, objItem.dValorPIS, _
                    objItem.dValorBaseCofins, objItem.dValorCofins, objItem.dValorPisRetido, objItem.dValorPisST, objItem.dValorCofinsRetido, _
                    objItem.dValorCofinsST, objItem.dValorISS, objItem.dPrecoUnitario, objItem.dPrecoTotal, objItem.sNatBCCred, objItem.dValorRecebido, objItem.dValorNoCR, objItem.sNCM, objItem.sCodAtividadeTab511, objItem.dAliquotaContrPrev, objItem.dValorRecContrPrev)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209401
                  
            lErro = objEFD.objTela.Processa_Item(False)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            objEFD.lNumTotalProcessados = objEFD.lNumTotalProcessados + 1
        
        Next
    End If

    SpedFiscal_Docs_Grava = SUCESSO

    Exit Function

Erro_SpedFiscal_Docs_Grava:

    SpedFiscal_Docs_Grava = gErr

    Select Case gErr
    
        Case 209400
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_SPEDDOCFISCAIS", gErr)
        
        Case 209401
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_SPEDDOCFISCAISITENS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209402)

    End Select

    Exit Function
    
End Function

Private Function SpedFiscal_Docs_Prepara(ByVal lComando As Long, ByVal lNumIntSped As Long, ByVal sBloco As String, tDoc As typeSpedDocFiscais) As Long

Dim lErro As Long

On Error GoTo Erro_SpedFiscal_Docs_Prepara

    'Alocação de espaço no buffer
    tDoc.sBloco = String(STRING_MAXIMO, 0)
    tDoc.sRegistro = String(STRING_MAXIMO, 0)
    tDoc.sOperacao = String(STRING_MAXIMO, 0)
    tDoc.sEmitente = String(STRING_MAXIMO, 0)
    tDoc.sModelo = String(STRING_MAXIMO, 0)
    tDoc.sSerie = String(STRING_MAXIMO, 0)
    tDoc.sChaveEletronica = String(STRING_MAXIMO, 0)
    tDoc.sFrete = String(STRING_MAXIMO, 0)
    tDoc.sPagamento = String(STRING_MAXIMO, 0)
    tDoc.sCgc = String(STRING_MAXIMO, 0)
    tDoc.sIndNatFrtPis = String(STRING_MAXIMO, 0)
    tDoc.sIndNatFrtCofins = String(STRING_MAXIMO, 0)

    'Le a tabelaSpedDocFiscais
    lErro = Comando_Executar(lComando, "SELECT NumIntSped, TipoDoc, NumIntDoc, Bloco, Registro, FilialEmpresa, " & _
                "Cliente, FilialCli, Fornecedor, FilialForn, DataEmissao, DataES, ValorTotal, " & _
                "Operacao, Emitente, Modelo, Situacao, Serie, NumDocumento, ChaveEletronica, " & _
                "Frete, Pagamento, ValorDesconto, ValorFrete, ValorSeguro, ValorDespesas, ValorServico, " & _
                "ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, ValorICMSST, ValorBaseIPI, ValorIPI, " & _
                "ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, ValorPisRetido, ValorPisST, ValorCofinsRetido, " & _
                "ValorCofinsST, ValorISS, CGC, CodMensagem, ValorMercadoria, IndNatFrtPis, IndNatFrtCofins, ValorRecebido, ValorNoCR, ValorRecContrPrev FROM SpedDocFiscais WHERE NumIntSped= ? AND Bloco= ? ORDER BY CGC, Registro ", _
                tDoc.lNumIntSped, tDoc.iTipoDoc, tDoc.lNumIntDoc, tDoc.sBloco, _
                tDoc.sRegistro, tDoc.iFilialEmpresa, tDoc.lCliente, tDoc.iFilialCli, tDoc.lFornecedor, _
                tDoc.iFilialForn, tDoc.dtDataEmissao, tDoc.dtDataES, tDoc.dValorTotal, tDoc.sOperacao, _
                tDoc.sEmitente, tDoc.sModelo, tDoc.iSituacao, tDoc.sSerie, tDoc.lNumDocumento, _
                tDoc.sChaveEletronica, tDoc.sFrete, tDoc.sPagamento, tDoc.dValorDesconto, tDoc.dValorFrete, _
                tDoc.dValorSeguro, tDoc.dValorDespesas, tDoc.dValorServico, tDoc.dValorNaoTributado, tDoc.dValorBaseICMS, _
                tDoc.dValorICMS, tDoc.dValorBaseICMSST, tDoc.dValorICMSST, tDoc.dValorBaseIPI, tDoc.dValorIPI, _
                tDoc.dValorBasePIS, tDoc.dValorPIS, tDoc.dValorBaseCofins, tDoc.dValorCofins, tDoc.dValorPisRetido, _
                tDoc.dValorPisST, tDoc.dValorCofinsRetido, tDoc.dValorCofinsST, tDoc.dValorISS, tDoc.sCgc, tDoc.lCodMensagem, tDoc.dValorMercadoria, tDoc.sIndNatFrtPis, tDoc.sIndNatFrtCofins, tDoc.dValorRecebido, tDoc.dValorNoCR, tDoc.dValorRecContrPrev, _
                lNumIntSped, sBloco)
    If lErro <> AD_SQL_SUCESSO Then gError 209403
    
    SpedFiscal_Docs_Prepara = SUCESSO

    Exit Function

Erro_SpedFiscal_Docs_Prepara:

    SpedFiscal_Docs_Prepara = gErr

    Select Case gErr

        Case 209403
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209404)

    End Select

    Exit Function

End Function
    
Public Function SpedFiscal_DocItens_Le(ByVal lComando As Long, ByVal objDoc As ClassSpedDocFiscais) As Long

Dim lErro As Long
Dim tItem As typeSpedDocFiscaisItens
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_SpedFiscal_DocItens_Le

    'Alocação de espaço no buffer
    tItem.sRegistro = String(STRING_MAXIMO, 0)
    tItem.sProduto = String(STRING_MAXIMO, 0)
    tItem.sDescricao = String(STRING_MAXIMO, 0)
    tItem.sUM = String(STRING_MAXIMO, 0)
    tItem.sMovFisica = String(STRING_MAXIMO, 0)
    tItem.sICMSCST = String(STRING_MAXIMO, 0)
    tItem.sCFOP = String(STRING_MAXIMO, 0)
    tItem.sIPICST = String(STRING_MAXIMO, 0)
    tItem.sIPICodEnq = String(STRING_MAXIMO, 0)
    tItem.sPISCST = String(STRING_MAXIMO, 0)
    tItem.sCofinsCST = String(STRING_MAXIMO, 0)
    tItem.sNatBCCred = String(STRING_MAXIMO, 0)
    tItem.sNCM = String(STRING_MAXIMO, 0)
    tItem.sCodAtividadeTab511 = String(STRING_MAXIMO, 0)

    'Le a tabelaSpedDocFiscaisItens
    lErro = Comando_Executar(lComando, "SELECT NumIntSped, TipoDoc, NumIntDoc, Item, Registro, Produto, " & _
                "Descricao, Quantidade, PrecoUnitario, PrecoTotal, UM, MovFisica, ICMSCST, CFOP, ICMSAliquota, " & _
                "ICMSSTAliquota, IPICST, IPICodEnq, IPIALiquota, PISCST, PISAliquota, PISBCQtd, " & _
                "PISAliquotaQtd, CofinsCST, CofinsAliquota, CofinsBCQtd, CofinsAliquotaQtd, ValorDesconto, ValorFrete, " & _
                "ValorSeguro, ValorDespesas, ValorServico, ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, " & _
                "ValorICMSST, ValorBaseIPI, ValorIPI, ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, " & _
                "ValorPisRetido, ValorPisST, ValorCofinsRetido, ValorCofinsST, ValorISS, NatBCCred, ValorRecebido, ValorNoCR, NCM, CodAtividadeTab511, AliquotaContrPrev, ValorRecContrPrev FROM SpedDocFiscaisItens WHERE NumIntSped= ?  AND TipoDoc= ?  AND NumIntDoc= ? ", _
                tItem.lNumIntSped, tItem.iTipoDoc, tItem.lNumIntDoc, tItem.iItem, _
                tItem.sRegistro, tItem.sProduto, tItem.sDescricao, tItem.dQuantidade, tItem.dPrecoUnitario, tItem.dPrecoTotal, tItem.sUM, _
                tItem.sMovFisica, tItem.sICMSCST, tItem.sCFOP, tItem.dICMSAliquota, tItem.dICMSSTAliquota, _
                tItem.sIPICST, tItem.sIPICodEnq, tItem.dIPIAliquota, tItem.sPISCST, tItem.dPISAliquota, _
                tItem.dPISBCQtd, tItem.dPISAliquotaQtd, tItem.sCofinsCST, tItem.dCOFINSAliquota, tItem.dCofinsBCQtd, _
                tItem.dCofinsAliquotaQtd, tItem.dValorDesconto, tItem.dValorFrete, tItem.dValorSeguro, tItem.dValorDespesas, _
                tItem.dValorServico, tItem.dValorNaoTributado, tItem.dValorBaseICMS, tItem.dValorICMS, tItem.dValorBaseICMSST, _
                tItem.dValorICMSST, tItem.dValorBaseIPI, tItem.dValorIPI, tItem.dValorBasePIS, tItem.dValorPIS, _
                tItem.dValorBaseCofins, tItem.dValorCofins, tItem.dValorPisRetido, tItem.dValorPisST, tItem.dValorCofinsRetido, _
                tItem.dValorCofinsST, tItem.dValorISS, tItem.sNatBCCred, tItem.dValorRecebido, tItem.dValorNoCR, tItem.sNCM, tItem.sCodAtividadeTab511, tItem.dAliquotaContrPrev, tItem.dValorRecContrPrev, _
                objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209405

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209406

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItem = New ClassSpedDocFiscaisItens
        
        Call SpedFiscal_DocItem_Copia(objItem, tItem)
        
        objDoc.colItens.Add objItem
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209407
        
    Loop

    SpedFiscal_DocItens_Le = SUCESSO

    Exit Function

Erro_SpedFiscal_DocItens_Le:

    SpedFiscal_DocItens_Le = gErr

    Select Case gErr

        Case 209405 To 209407
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAISITENS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209408)

    End Select

    Exit Function

End Function

Private Function SpedFiscal_Doc_Copia(objDoc As ClassSpedDocFiscais, tDoc As typeSpedDocFiscais) As Long
On Error GoTo Erro_SpedFiscal_Doc_Copia
    objDoc.lNumIntSped = tDoc.lNumIntSped
    objDoc.iTipoDoc = tDoc.iTipoDoc
    objDoc.lNumIntDoc = tDoc.lNumIntDoc
    objDoc.sBloco = tDoc.sBloco
    objDoc.sRegistro = tDoc.sRegistro
    objDoc.iFilialEmpresa = tDoc.iFilialEmpresa
    objDoc.lCliente = tDoc.lCliente
    objDoc.iFilialCli = tDoc.iFilialCli
    objDoc.lFornecedor = tDoc.lFornecedor
    objDoc.iFilialForn = tDoc.iFilialForn
    objDoc.dtDataEmissao = tDoc.dtDataEmissao
    objDoc.dtDataES = tDoc.dtDataES
    objDoc.dValorTotal = tDoc.dValorTotal
    objDoc.sOperacao = tDoc.sOperacao
    objDoc.sEmitente = tDoc.sEmitente
    objDoc.sModelo = tDoc.sModelo
    objDoc.iSituacao = tDoc.iSituacao
    objDoc.sSerie = tDoc.sSerie
    objDoc.lNumDocumento = tDoc.lNumDocumento
    objDoc.sChaveEletronica = tDoc.sChaveEletronica
    objDoc.sFrete = tDoc.sFrete
    objDoc.sPagamento = tDoc.sPagamento
    objDoc.dValorDesconto = tDoc.dValorDesconto
    objDoc.dValorFrete = tDoc.dValorFrete
    objDoc.dValorSeguro = tDoc.dValorSeguro
    objDoc.dValorDespesas = tDoc.dValorDespesas
    objDoc.dValorServico = tDoc.dValorServico
    objDoc.dValorNaoTributado = tDoc.dValorNaoTributado
    objDoc.dValorBaseICMS = tDoc.dValorBaseICMS
    objDoc.dValorICMS = tDoc.dValorICMS
    objDoc.dValorBaseICMSST = tDoc.dValorBaseICMSST
    objDoc.dValorICMSST = tDoc.dValorICMSST
    objDoc.dValorBaseIPI = tDoc.dValorBaseIPI
    objDoc.dValorIPI = tDoc.dValorIPI
    objDoc.dValorBasePIS = tDoc.dValorBasePIS
    objDoc.dValorPIS = tDoc.dValorPIS
    objDoc.dValorBaseCofins = tDoc.dValorBaseCofins
    objDoc.dValorCofins = tDoc.dValorCofins
    objDoc.dValorPisRetido = tDoc.dValorPisRetido
    objDoc.dValorPisST = tDoc.dValorPisST
    objDoc.dValorCofinsRetido = tDoc.dValorCofinsRetido
    objDoc.dValorCofinsST = tDoc.dValorCofinsST
    objDoc.dValorISS = tDoc.dValorISS
    objDoc.sCgc = tDoc.sCgc
    objDoc.lCodMensagem = tDoc.lCodMensagem
    objDoc.dValorMercadoria = tDoc.dValorMercadoria
    objDoc.sIndNatFrtPis = tDoc.sIndNatFrtPis
    objDoc.sIndNatFrtCofins = tDoc.sIndNatFrtCofins
    objDoc.dValorRecebido = tDoc.dValorRecebido
    objDoc.dValorNoCR = tDoc.dValorNoCR
    objDoc.dValorRecContrPrev = tDoc.dValorRecContrPrev
    SpedFiscal_Doc_Copia = SUCESSO
    Exit Function
Erro_SpedFiscal_Doc_Copia:
    SpedFiscal_Doc_Copia = gErr
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209409)
    End Select
    Exit Function
End Function

Private Function SpedFiscal_DocItem_Copia(objItem As ClassSpedDocFiscaisItens, tItem As typeSpedDocFiscaisItens) As Long
On Error GoTo Erro_SpedFiscal_DocItem_Copia
    objItem.lNumIntSped = tItem.lNumIntSped
    objItem.iTipoDoc = tItem.iTipoDoc
    objItem.lNumIntDoc = tItem.lNumIntDoc
    objItem.iItem = tItem.iItem
    objItem.sRegistro = tItem.sRegistro
    objItem.sProduto = tItem.sProduto
    objItem.sDescricao = tItem.sDescricao
    objItem.dQuantidade = tItem.dQuantidade
    objItem.sUM = tItem.sUM
    objItem.sMovFisica = tItem.sMovFisica
    objItem.sICMSCST = tItem.sICMSCST
    objItem.sCFOP = tItem.sCFOP
    objItem.dICMSAliquota = tItem.dICMSAliquota
    objItem.dICMSSTAliquota = tItem.dICMSSTAliquota
    objItem.sIPICST = tItem.sIPICST
    objItem.sIPICodEnq = tItem.sIPICodEnq
    objItem.dIPIAliquota = tItem.dIPIAliquota
    objItem.sPISCST = tItem.sPISCST
    objItem.dPISAliquota = tItem.dPISAliquota
    objItem.dPISBCQtd = tItem.dPISBCQtd
    objItem.dPISAliquotaQtd = tItem.dPISAliquotaQtd
    objItem.sCofinsCST = tItem.sCofinsCST
    objItem.dCOFINSAliquota = tItem.dCOFINSAliquota
    objItem.dCofinsBCQtd = tItem.dCofinsBCQtd
    objItem.dCofinsAliquotaQtd = tItem.dCofinsAliquotaQtd
    objItem.dValorDesconto = tItem.dValorDesconto
    objItem.dValorFrete = tItem.dValorFrete
    objItem.dValorSeguro = tItem.dValorSeguro
    objItem.dValorDespesas = tItem.dValorDespesas
    objItem.dValorServico = tItem.dValorServico
    objItem.dValorNaoTributado = tItem.dValorNaoTributado
    objItem.dValorBaseICMS = tItem.dValorBaseICMS
    objItem.dValorICMS = tItem.dValorICMS
    objItem.dValorBaseICMSST = tItem.dValorBaseICMSST
    objItem.dValorICMSST = tItem.dValorICMSST
    objItem.dValorBaseIPI = tItem.dValorBaseIPI
    objItem.dValorIPI = tItem.dValorIPI
    objItem.dValorBasePIS = tItem.dValorBasePIS
    objItem.dValorPIS = tItem.dValorPIS
    objItem.dValorBaseCofins = tItem.dValorBaseCofins
    objItem.dValorCofins = tItem.dValorCofins
    objItem.dValorPisRetido = tItem.dValorPisRetido
    objItem.dValorPisST = tItem.dValorPisST
    objItem.dValorCofinsRetido = tItem.dValorCofinsRetido
    objItem.dValorCofinsST = tItem.dValorCofinsST
    objItem.dValorISS = tItem.dValorISS
    objItem.dPrecoUnitario = tItem.dPrecoUnitario
    objItem.dPrecoTotal = tItem.dPrecoTotal
    objItem.sNatBCCred = tItem.sNatBCCred
    objItem.dValorRecebido = tItem.dValorRecebido
    objItem.dValorNoCR = tItem.dValorNoCR
    objItem.sNCM = tItem.sNCM
    objItem.sCodAtividadeTab511 = tItem.sCodAtividadeTab511
    objItem.dAliquotaContrPrev = tItem.dAliquotaContrPrev
    objItem.dValorRecContrPrev = tItem.dValorRecContrPrev
    SpedFiscal_DocItem_Copia = SUCESSO
    Exit Function
Erro_SpedFiscal_DocItem_Copia:
    SpedFiscal_DocItem_Copia = gErr
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209410)
    End Select
    Exit Function
End Function

Function SpedFiscal_Gravar_Arq(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long, lErro1 As Long
Dim sRegistro As String, lSeq As Long
Dim sRegistroAux As String, lSeqAnt As Long
Dim lDifQtdRegs As Long, lIndice As Long

On Error GoTo Erro_SpedFiscal_Gravar_Arq

    'A previsão de falta é o que estava inicialmente previsto menos o que já foi processado
    'O real de falta é q quantidade total de registro que vai ser colocada no arquivo texto
    lDifQtdRegs = (objEFD.lNumTotalPrevisto - objEFD.lNumTotalProcessados) - objEFD.lNumTotalReal
    'Se estava previsto mais registros do que realmente falta avança a barra de progressão
    If lDifQtdRegs > 0 Then
        For lIndice = 1 To lDifQtdRegs
            lErro = objEFD.objTela.Processa_Item(False) 'False significa que não está processando o item, só processa quando gera o registro no BD
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Next
    End If

    Open objEFD.sDiretorio For Output As #1

    sRegistroAux = String(STRING_MAXIMO, 0)
    sRegistro = ""

    lErro = Comando_Executar(objEFD.ProxComando, "SELECT R.Seq, R.Registro FROM SpedFiscalPis AS S, SpedFiscalPisReg AS R WHERE S.NumIntSped = R.NumIntSped AND S.Seq = R.Seq AND S.NumIntSped = ? ORDER BY R.Seq, R.SeqReg ", _
    lSeq, sRegistroAux, objEFD.lNumIntSped)
    If lErro <> AD_SQL_SUCESSO Then gError 209411

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209412
    
    lSeqAnt = lSeq

    Do While lErro = AD_SQL_SUCESSO
    
        sRegistro = sRegistro + sRegistroAux

        lErro = Comando_BuscarProximo(objEFD.UltimoCmdUsado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209413

        If lSeqAnt <> lSeq Or lErro = AD_SQL_SEM_DADOS Then
            lSeqAnt = lSeq
            Print #1, sRegistro
            If lDifQtdRegs >= 0 Then
                lErro1 = objEFD.objTela.Processa_Item(False)
                If lErro1 <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Else
                'Se falta mais do que estava previsto não mexe na barra de progressão até falta igual
                lDifQtdRegs = lDifQtdRegs + 1
            End If
            sRegistro = ""
        End If

    Loop

    Close #1

    SpedFiscal_Gravar_Arq = SUCESSO

    Exit Function

Erro_SpedFiscal_Gravar_Arq:

    SpedFiscal_Gravar_Arq = gErr

   Select Case gErr
        
        Case 209411, 209412, 209413
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209414)

    End Select

    Close #1

End Function

Private Function SpedFiscal_Trata_ValorRecebido(ByVal objDoc As ClassSpedDocFiscais, ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim dValorNoCRItens1 As Double
Dim dValorNoCRItens2 As Double
Dim dValorRecebidoItens As Double
Dim dDif As Double, dFator As Double, dSoma As Double

On Error GoTo Erro_SpedFiscal_Trata_ValorRecebido

    'Rateia o valor a receber por item para poder quebrar por CST caso tenham CSTs diferentes
    dValorNoCRItens1 = 0
    For Each objItem In objDoc.colItens
        objItem.dValorNoCR = objItem.dPrecoTotal + objItem.dValorFrete + objItem.dValorDespesas + objItem.dValorSeguro + objItem.dValorICMSST + objItem.dValorIPI  'Teria que ratear IR, ver se o item é cobrado, etc, mas para efeito de cálculo inicial será considerado apenas valores básicos e depois será ajustado de acordo com o a receber gerado
        dValorNoCRItens1 = dValorNoCRItens1 + objItem.dValorNoCR
    Next
    dValorNoCRItens2 = dValorNoCRItens1
    
    If Abs(objDoc.dValorNoCR - dValorNoCRItens1) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorNoCR - dValorNoCRItens1
        dValorNoCRItens2 = 0
        For Each objItem In objDoc.colItens
            objItem.dValorNoCR = objItem.dValorNoCR + Arredonda_Moeda(objItem.dValorNoCR * (dDif / dValorNoCRItens1))
            dValorNoCRItens2 = dValorNoCRItens2 + objItem.dValorNoCR
        Next
    End If

    If Abs(objDoc.dValorNoCR - dValorNoCRItens2) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorNoCR - dValorNoCRItens2
        For Each objItem In objDoc.colItens
            objItem.dValorNoCR = objItem.dValorNoCR + dDif
            Exit For
        Next
    End If
    
    'Distribui o valor recebido por item
    dValorRecebidoItens = 0
    dFator = objDoc.dValorRecebido / objDoc.dValorNoCR
    For Each objItem In objDoc.colItens
        objItem.dValorRecebido = Arredonda_Moeda(objItem.dValorNoCR * dFator)
        dValorRecebidoItens = dValorRecebidoItens + objItem.dValorRecebido
    Next
    
    If Abs(objDoc.dValorRecebido - dValorRecebidoItens) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorRecebido - dValorRecebidoItens
        For Each objItem In objDoc.colItens
            objItem.dValorRecebido = objItem.dValorRecebido + dDif
            Exit For
        Next
    End If

    'Calcula o valor recebido sujeito a contribuição previdenciária
    dSoma = 0
    For Each objItem In objDoc.colItens
        If Len(Trim(objItem.sCodAtividadeTab511)) > 0 Then
            objItem.dValorRecContrPrev = objItem.dValorRecebido
            dSoma = dSoma + objItem.dValorRecContrPrev
        End If
    Next
    objDoc.dValorRecContrPrev = dSoma
    
    SpedFiscal_Trata_ValorRecebido = SUCESSO

    Exit Function

Erro_SpedFiscal_Trata_ValorRecebido:

    SpedFiscal_Trata_ValorRecebido = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209402)

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

Function SFF_Data(ByVal dtData As Date) As String
    If dtData <> DATA_NULA Then
        SFF_Data = Format(dtData, "ddmmyyyy")
    Else
        SFF_Data = ""
    End If
End Function

Function SFF_Num(ByVal dNum As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Num = Format(dNum, sFormato)
End Function

Function SFF_Perc(ByVal dPerc As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Perc = Format(dPerc * 100, sFormato)
End Function

Private Function Processa_Registro_0000(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA
'JURÍDICA
Dim lErro As Long
Dim objCidades As New ClassCidades
Dim objSpedRegistro As ClassSpedRegistro
Dim sRazaoSocial As String
Dim lComandoDic As Long
Dim sIE As String, sIND_SIT_ESP As String, sIND_ATIV As String
Dim lComando As Long, sRazaoSocialAux1 As String, sRazaoSocialAux2 As String

On Error GoTo Erro_Processa_Registro_0000

    '******** registro 0000 *****************
    objEFD.objSpedFiscal.sTipo = "0000"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    'abre conexao com o BD
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 209415

    lComando = objEFD.ProxComando
    
    sRazaoSocialAux1 = String(STRING_NOME_EMPRESA, 0)
    
    'le a empresa
    lErro = Comando_Executar(lComandoDic, "SELECT Nome FROM Empresas WHERE Codigo = ? ", sRazaoSocialAux1, glEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209416
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209417
       
    objCidades.sDescricao = objEFD.objFilialEmpresa.objEndereco.sCidade
    
    sRazaoSocial = sRazaoSocialAux1
    
    sRazaoSocialAux2 = String(255, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM AdmConfig WHERE Codigo = ? AND FilialEmpresa = ? ", sRazaoSocialAux2, "EMPRESA_RAZAO_SOCIAL", objEFD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209416
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209417
    
    If lErro = AD_SQL_SUCESSO Then sRazaoSocial = sRazaoSocialAux2
    
    lErro = CF("Cidade_Le_Nome", objCidades)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(objCidades.sCodIBGE)) = 0 Then gError 209418
    
    Call Formata_String_Numero(objEFD.objFilialEmpresa.sInscricaoEstadual, sIE)
    
    If objEFD.iSituacaoEspecial <> -1 Then
        sIND_SIT_ESP = CStr(objEFD.iSituacaoEspecial)
    Else
        sIND_SIT_ESP = ""
    End If
    
    'iTipoTribAtividade
    'AtivTrib(1) = Comercial
    'AtivTrib(2) = Industrial
    'AtivTrib(3) = Prestação de Serviços
    'IND_ATIV
    '0 – Industrial ou equiparado a industrial;
    '1 – Prestador de serviços;
    '2 - Atividade de comércio;
    '3 – Pessoas jurídicas referidas nos §§ 6º, 8º e 9º do art. 3º
    'da Lei nº 9.718, de 1998;
    '4 – Atividade imobiliária;
    '9 – Outros
    Select Case objEFD.objFilialEmpresa.iTipoTribAtividade
        Case 1
            sIND_ATIV = "2"
        Case 2
            sIND_ATIV = "0"
        Case 3
            sIND_ATIV = "1"
        Case Else
            sIND_ATIV = "9"
    End Select
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "0000")
    Call objEFD.Reg_Campo_add(2, "COD_VER", Format(objEFD.iCodVersao, "000"))
    Call objEFD.Reg_Campo_add(3, "TIPO_ESCRIT", CStr(objEFD.iTipo))
    Call objEFD.Reg_Campo_add(4, "IND_SIT_ESP", sIND_SIT_ESP)
    Call objEFD.Reg_Campo_add(5, "NUM_REC_ANTERIOR", SFF_Texto(objEFD.sReciboAnterior))
    Call objEFD.Reg_Campo_add(6, "DT_INI", SFF_Data(objEFD.dtDataIni))
    Call objEFD.Reg_Campo_add(7, "DT_FIN", SFF_Data(objEFD.dtDataFim))
    Call objEFD.Reg_Campo_add(8, "NOME", SFF_Texto(sRazaoSocial, 100, , , True))
    Call objEFD.Reg_Campo_add(9, "CNPJ", IIf(Len(objEFD.objFilialEmpresa.sCgc) = 14, objEFD.objFilialEmpresa.sCgc, ""))
    Call objEFD.Reg_Campo_add(10, "UF", objEFD.objFilialEmpresa.objEndereco.sSiglaEstado)
    Call objEFD.Reg_Campo_add(11, "COD_MUN", objCidades.sCodIBGE)
    Call objEFD.Reg_Campo_add(12, "SUFRAMA", "")
    Call objEFD.Reg_Campo_add(13, "IND_NAT_PJ", "00")
    Call objEFD.Reg_Campo_add(14, "IND_ATIV", sIND_ATIV)
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Comando_Fechar(lComandoDic)

    Processa_Registro_0000 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0000:

   Processa_Registro_0000 = gErr

   Select Case gErr
    
        Case 209415
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 209416, 209417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESAS", gErr)
            
        Case 209418
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO2", gErr, objCidades.sDescricao)
           
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209419)
     
    End Select

    Call Comando_Fechar(lComandoDic)

    Exit Function
    
End Function

Private Function Processa_Registro_0001(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0001: ABERTURA DO BLOCO 0
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_0001

    '******** registro 0001 *****************
    objEFD.objSpedFiscal.sTipo = "0001"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "0001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", "0")
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0001:

   Processa_Registro_0001 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209420)
    
    End Select
    
    Exit Function

End Function
            
Private Function Processa_Registro_0100(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0100: DADOS DO CONTABILISTA
Dim lErro As Long
Dim stelefone As String
Dim sFax As String
Dim objCidades As New ClassCidades

On Error GoTo Erro_Processa_Registro_0100
    
    '******** registro 0100 *****************
    objEFD.objSpedFiscal.sTipo = "0100"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    If Len(Trim(objEFD.objFilialEmpresa.sContador)) > 0 Then
    
        If Len(Trim(objEFD.objFilialEmpresa.sCPFContador)) = 0 Then gError 209500
        If Len(Trim(objEFD.objFilialEmpresa.sCRCContador)) = 0 Then gError 209501
    
        Call Formata_String_Numero(objEFD.objFilialEmpresa.objEnderecoContador.sTelefone1, stelefone)
        Call Formata_String_Numero(objEFD.objFilialEmpresa.objEnderecoContador.sFax, sFax)
        
        objCidades.sDescricao = objEFD.objFilialEmpresa.objEnderecoContador.sCidade
        
        If Len(Trim(objCidades.sDescricao)) > 0 Then
        
            lErro = CF("Cidade_Le_Nome", objCidades)
            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "0100")
        Call objEFD.Reg_Campo_add(2, "NOME", SFF_Texto(objEFD.objFilialEmpresa.sContador, 100))
        Call objEFD.Reg_Campo_add(3, "CPF", SFF_Texto(objEFD.objFilialEmpresa.sCPFContador, 11))
        Call objEFD.Reg_Campo_add(4, "CRC", SFF_Texto(objEFD.objFilialEmpresa.sCRCContador, 15))
        Call objEFD.Reg_Campo_add(5, "CNPJ", SFF_Texto(objEFD.objFilialEmpresa.sCNPJContador, 14))
        Call objEFD.Reg_Campo_add(6, "CEP", SFF_Texto(objEFD.objFilialEmpresa.objEnderecoContador.sCEP))
        Call objEFD.Reg_Campo_add(7, "END", SFF_Texto(objEFD.objFilialEmpresa.objEnderecoContador.sLogradouro, 60))
        Call objEFD.Reg_Campo_add(8, "NUM", CStr(objEFD.objFilialEmpresa.objEnderecoContador.lNumero))
        Call objEFD.Reg_Campo_add(9, "COMPL", SFF_Texto(objEFD.objFilialEmpresa.objEnderecoContador.sComplemento, 60))
        Call objEFD.Reg_Campo_add(10, "BAIRRO", SFF_Texto(objEFD.objFilialEmpresa.objEnderecoContador.sBairro, 60))
        Call objEFD.Reg_Campo_add(11, "FONE", SFF_Texto(stelefone, 10))
        Call objEFD.Reg_Campo_add(12, "FAX", SFF_Texto(sFax, 10))
        Call objEFD.Reg_Campo_add(13, "EMAIL", SFF_Texto(objEFD.objFilialEmpresa.objEnderecoContador.sEmail))
        Call objEFD.Reg_Campo_add(14, "COD_MUN", objCidades.sCodIBGE)
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Processa_Registro_0100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0100:

   Processa_Registro_0100 = gErr

   Select Case gErr

        Case 209500
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTADOR_SEM_CPF", gErr)

        Case 209501
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTADOR_SEM_CRC", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209421)
     
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0110(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0110: REGIMES DE APURAÇÃO DA CONTRIBUIÇÃO SOCIAL E DE
'APROPRIAÇÃO DE CRÉDITO
'Este registro tem por objetivo definir o regime de incidência a que se submete a pessoa jurídica (não-cumulativo,
'cumulativo ou ambos os regimes) no período da escrituração. No caso de sujeição ao regime não-cumulativo, será
'informado também o método de apropriação do crédito incidente sobre operações comuns a mais de um tipo de receita
'adotado pela pessoa jurídica para o ano-calendário.
Dim lErro As Long
Dim sCOD_INC_TRIB As String
Dim sIND_APRO_CRED As String
Dim sCOD_TIPO_CONT As String
Dim sIND_REG_CUM As String
Dim lComando As Long, lAux As Long, sChave As String
Dim sSQL As String

On Error GoTo Erro_Processa_Registro_0110

    '******** registro 0110 *****************
    objEFD.objSpedFiscal.sTipo = "0110"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    lComando = objEFD.ProxComando
    
    If objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = MARCADO Then
    
        'Código indicador da incidência tributária no período:
        '1 – Escrituração de operações com incidência
        'exclusivamente no regime não-cumulativo;
        '2 – Escrituração de operações com incidência
        'exclusivamente no regime cumulativo;
        '3 – Escrituração de operações com incidência nos
        'regimes não-cumulativo e cumulativo.
        
        sSQL = "SELECT D.NumIntSped"
        sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
        sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Registro = 'F500' AND PisAliquota <> 0 AND ABS(PisAliquota - 0.0065) < 0.0001 AND I.NumIntSped = ?"
    
        lErro = Comando_Executar(lComando, sSQL, lAux, objEFD.lNumIntSped)
        If lErro <> SUCESSO Then gError 211975
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211976
        
        If lErro <> AD_SQL_SUCESSO Then
            sCOD_INC_TRIB = "1"
        Else
            sCOD_INC_TRIB = "3"
        End If
        
        If objEFD.objFilialEmpresa.iPisCofinsIndAproCred <> 2 Then
            sIND_APRO_CRED = "1"
        Else
            sIND_APRO_CRED = "2"
        End If
        sCOD_TIPO_CONT = "1"
    Else
        sCOD_INC_TRIB = "2"
        sIND_APRO_CRED = ""
        sCOD_TIPO_CONT = ""
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "0110")
    Call objEFD.Reg_Campo_add(2, "COD_INC_TRIB", sCOD_INC_TRIB)
    Call objEFD.Reg_Campo_add(3, "IND_APRO_CRED", sIND_APRO_CRED)
    Call objEFD.Reg_Campo_add(4, "COD_TIPO_CONT", sCOD_TIPO_CONT)
    
    '2. O campo 05 (IND_REC_CUM) não deverá constar no arquivo da escrituração a ser importado pelo PVA,
    'versão 2.00, versão essa que acresce em relação à versão anterior (1.07), os registros da escrituração da
    'Contribuição Previdenciária sobre a Receita Bruta (Bloco P). O referido campo 05 só deverá constar no arquivo
    'da escrituração a ser importado pelo PVA na versão 2.01A, com previsão de disponibilização em julho de 2012,
    'que irá então constar com os registros da escrituração do PIS/Pasep e da Cofins, para a pessoa jurídica tributada
    'com base no lucro presumido.
    'Desta forma, no arquivo gerado para ser importado nas versões 1.07 e 2.00 do PVA, o registro “0110” deverá ser
    'informado com apenas 04 (quatro) campos.
    
    If objEFD.iCodVersao >= 3 Then
        If sCOD_INC_TRIB = "2" And objEFD.objFilialEmpresa.iLucroPresumido = MARCADO Then
            sIND_REG_CUM = CStr(objEFD.objFilialEmpresa.iPisCofinsRegCumTipo)
        Else
            sIND_REG_CUM = ""
        End If
        Call objEFD.Reg_Campo_add(5, "IND_REG_CUM", sIND_REG_CUM)
    End If

    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If sIND_APRO_CRED = "2" Then
    
        'REGISTRO 0111: TABELA DE RECEITA BRUTA MENSAL PARA FINS DE RATEIO DE CRÉDITOS COMUNS
        Call objEFD.objTela.SPED_Status("Processando Registro 0111", "")
        lErro = Processa_Registro_0111(objEFD, objEFD.objSpedFiscal.lSeq, objEFD.ProxComando)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Processa_Registro_0110 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0110:

   Processa_Registro_0110 = gErr

   Select Case gErr
   
        Case 211975 To 211976
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209422)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0110", sChave)
    
    Exit Function

End Function

Private Function Processa_Registro_0111(ByVal objEFD As ClassEFDPisCofinsSel, ByVal lSeqPai As Long, ByVal lComando As Long) As Long
'REGISTRO 0111: TABELA DE RECEITA BRUTA MENSAL PARA FINS DE RATEIO DE
'CRÉDITOS COMUNS
'Este registro é de preenchimento obrigatório, sempre que for informado no Registro “0110”, Campo 03
'(IND_APRO_CRED), o indicador correspondente ao método do Rateio Proporcional com base na Receita Bruta
'(indicador “2”), na apuração de créditos vinculados a mais de um tipo de receita.
'De acordo com a legislação que instituiu a não cumulatividade da Contribuição para o PIS/Pasep (Lei nº 10.637/02, art.
'1º, § 1º) e da Cofins (Lei nº 10.833/03, art. 1º, § 1º), a Receita Bruta compreende a receita da venda de bens e serviços nas
'operações em conta própria ou alheia (comissões pela intermediação de negócios).
'No tocante às receitas de natureza cumulativa, considera-se como Receita Bruta, como definida pela legislação do
'imposto de renda, a proveniente da venda de bens nas operações de conta própria, do preço dos serviços prestados e do
'resultado auferido nas operações de conta alheia (Lei nº 9.715/98, art. 3º e Decreto-Lei nº 1.598/77, art. 12).
'Assim, de acordo com a legislação tributária e os princípios contábeis básicos, as receitas diversas que não sejam
'decorrentes da venda de bens e serviços nas operações em conta própria ou alheia, não se classificam como receita bruta,
'não devendo desta forma ser consideradas para fins de rateio no registro “0111”.
'A título exemplificativo, uma empresa que tenha por objeto social a fabricação de bens (industria) ou a revenda de bens
'(comércio), não devem considerar como receita bruta, para fins de rateio, por não serem classificadas como tal, entre
'outras:
'- as receitas não operacionais, decorrentes da venda de ativo imobilizado;
'- as receitas não próprias da atividade, de natureza financeira ou não, de aluguéis de bens móveis e imóveis, etc.;
'- de reversões de provisões e recuperações de créditos baixados como perda, que não representem ingresso de
'novas receitas;
'- do resultado positivo da avaliação de investimentos pelo valor do patrimônio líquido e os lucros e dividendos
'derivados de investimentos avaliados pelo custo de aquisição, que tenham sido computados como receita.

Dim lErro As Long, iIndice As Integer
Dim sChave As String
Dim sSQL As String
Dim dValorRecIntNaoCum As Double, dValorRecIntIsenta As Double
Dim dValorRecExp As Double, dValorRecIntCum As Double, dValorRec As Double

On Error GoTo Erro_Processa_Registro_0111

    '******************* registro 0111 **********************************
    objEFD.objSpedFiscal.sTipo = "0111"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai

    sSQL = "SELECT "
    sSQL = sSQL + " " + "SUM(CASE WHEN LEFT(CFOP,1) <> '7' AND PisAliquota <> 0 AND ABS(PisAliquota - 0.0065) >= 0.0001 THEN I.ValorRecebido ELSE 0 END) AS ValorRecIntNaoCum, "
    sSQL = sSQL + " " + "SUM(CASE WHEN LEFT(CFOP,1) <> '7' AND PisAliquota = 0 THEN I.ValorRecebido ELSE 0 END) AS ValorRecIntIsenta, "
    sSQL = sSQL + " " + "SUM(CASE WHEN LEFT(CFOP,1) = '7' THEN I.ValorRecebido ELSE 0 END) AS ValorRecExp, "
    sSQL = sSQL + " " + "SUM(CASE WHEN LEFT(CFOP,1) <> '7' AND ABS(PisAliquota - 0.0065) < 0.0001 THEN I.ValorRecebido ELSE 0 END) AS ValorRecIntCum, "
    sSQL = sSQL + " " + "SUM(I.ValorRecebido) AS ValorRec "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Registro = 'F500' AND I.NumIntSped = ?"

    lErro = Comando_Executar(lComando, sSQL, dValorRecIntNaoCum, dValorRecIntIsenta, dValorRecExp, dValorRecIntCum, dValorRec, objEFD.lNumIntSped)
    If lErro <> SUCESSO Then gError 211975

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211976
    
    sChave = ""
    Call objEFD.objTela.SPED_Status("Processando Registro 0111", sChave)

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "0111")
    Call objEFD.Reg_Campo_add(2, "REC_BRU_NCUM_TRIB_MI", SFF_Num(dValorRecIntNaoCum))
    Call objEFD.Reg_Campo_add(3, "REC_BRU_ NCUM_NT_MI", SFF_Num(dValorRecIntIsenta))
    Call objEFD.Reg_Campo_add(4, "REC_BRU_ NCUM_EXP", SFF_Num(dValorRecExp))
    Call objEFD.Reg_Campo_add(5, "REC_BRU_CUM", SFF_Num(dValorRecIntCum))
    Call objEFD.Reg_Campo_add(6, "REC_BRU_TOTAL", SFF_Num(dValorRec))
    Call objEFD.Finaliza_Reg

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0111 = SUCESSO

    Exit Function

Erro_Processa_Registro_0111:

   Processa_Registro_0111 = gErr

   Select Case gErr
   
        Case 211975 To 211977
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211978)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0111", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_0140(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0140: TABELA DE CADASTRO DE ESTABELECIMENTO
'Este registro tem por objetivo relacionar e informar os estabelecimentos da pessoa jurídica que auferiram receitas no
'período da escrituração, que realizaram operações com direito a créditos ou que sofreram retenções na fonte, no período
'da escrituração.
Dim lErro As Long
Dim objCidades As ClassCidades
Dim sIM As String, sIE As String
Dim objFilialEmpresa As AdmFiliais, lComando As Long, iFilialEmpresa As Integer

On Error GoTo Erro_Processa_Registro_0140

    lComando = objEFD.ProxComando
    
    'lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM FiliaisEmpresa WHERE FilialEmpresa=1 OR FilialEmpresa IN (SELECT DISTINCT FilialEmpresa FROM SpedDocFiscais WHERE NumIntSped = ?) AND FilialEmpresa < 50 ORDER BY FilialEmpresa", iFilialEmpresa, objEFD.lNumIntSped)
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM FiliaisEmpresa WHERE FilialEmpresa IN (SELECT DISTINCT FilialEmpresa FROM SpedDocFiscais WHERE NumIntSped = ?) AND FilialEmpresa < 50 ORDER BY FilialEmpresa", iFilialEmpresa, objEFD.lNumIntSped)
    If lErro <> AD_SQL_SUCESSO Then gError 201291
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201292
            
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objFilialEmpresa = New AdmFiliais
    
        objFilialEmpresa.iCodFilial = iFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
        objEFD.objSpedFiscal.lSeq = objEFD.objSpedFiscal.lSeq + 1
    
        Set objCidades = New ClassCidades

        objCidades.sDescricao = objFilialEmpresa.objEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(objCidades.sCodIBGE)) = 0 Then gError 209424
        
        sIM = ""
        Call Formata_String_Numero(objFilialEmpresa.sInscricaoMunicipal, sIM)
        
        sIE = ""
        Call Formata_String_Numero(objFilialEmpresa.sInscricaoEstadual, sIE)
        
        '******** registro 0140 *****************
        objEFD.objSpedFiscal.sTipo = "0140"
        objEFD.objSpedFiscal.lSeqPai = 0
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "0140")
        Call objEFD.Reg_Campo_add(2, "COD_EST", Format(iFilialEmpresa, "00000"))
        Call objEFD.Reg_Campo_add(3, "NOME", SFF_Texto(objFilialEmpresa.sNome, 100))
        Call objEFD.Reg_Campo_add(4, "CNPJ", IIf(Len(objFilialEmpresa.sCgc) = 14, objFilialEmpresa.sCgc, ""))
        Call objEFD.Reg_Campo_add(5, "UF", objFilialEmpresa.objEndereco.sSiglaEstado)
        Call objEFD.Reg_Campo_add(6, "IE", SFF_Texto(sIE))
        Call objEFD.Reg_Campo_add(7, "COD_MUN", objCidades.sCodIBGE)
        Call objEFD.Reg_Campo_add(8, "IM", sIM)
        Call objEFD.Reg_Campo_add(9, "SUFRAMA", "")
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 0145: REGIME DE APURAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA
        If objFilialEmpresa.iContribPrevPorRecBruta = MARCADO Then
            Call objEFD.objTela.SPED_Status("Processando Registro 0145", "")
            lErro = Processa_Registro_0145(objEFD, objFilialEmpresa.iCodFilial)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        'REGISTRO 0150: TABELA DE CADASTRO DO PARTICIPANTE
        Call objEFD.objTela.SPED_Status("Processando Registro 0150", "")
        lErro = Processa_Registro_0150(objEFD, objFilialEmpresa.iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
        'REGISTRO 0190: IDENTIFICAÇÃO DAS UNIDADES DE MEDIDA
        Call objEFD.objTela.SPED_Status("Processando Registro 0190", "")
        lErro = Processa_Registro_0190(objEFD, objFilialEmpresa.iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTOS E SERVIÇOS)
        'REGISTRO 0205: ALTERAÇÃO DO ITEM
        'REGISTRO 0206: CÓDIGO DE PRODUTO CONFORME TABELA ANP (COMBUSTÍVEIS)
        'REGISTRO 0208: CÓDIGO DE GRUPOS POR MARCA COMERCIAL – REFRI (BEBIDAS FRIAS).
        Call objEFD.objTela.SPED_Status("Processando Registro 0200, 205, 206 e 208", "")
        lErro = Processa_Registro_0200(objEFD, objFilialEmpresa.iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 0400: TABELA DE NATUREZA DA OPERAÇÃO/PRESTAÇÃO
        Call objEFD.objTela.SPED_Status("Processando Registro 0400", "")
        lErro = Processa_Registro_0400(objEFD, objFilialEmpresa.iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'REGISTRO 0450: TABELA DE INFORMAÇÃO COMPLEMENTAR DO DOCUMENTO FISCAL
        Call objEFD.objTela.SPED_Status("Processando Registro 0450", "")
        lErro = Processa_Registro_0450(objEFD, objFilialEmpresa.iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201293

    Loop
    
    Processa_Registro_0140 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0140:

   Processa_Registro_0140 = gErr

   Select Case gErr
        
        Case 209424
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO2", gErr, objCidades.sDescricao)
           
        Case ERRO_SEM_MENSAGEM
        
        Case 201291, 201292, 201293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209425)
     
    End Select

    Exit Function
    
End Function

Private Function Processa_Registro_0145(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0145: REGIME DE APURAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA
'SOBRE A RECEITA BRUTA
'Este registro servirá para identificar a obrigatoriedade de escrituração da Contribuição Previdenciária sobre Receitas, no
'Bloco “P”. O Registro “0145” tem natureza meramente informativa, não transferindo nem recebendo valores de
'quaisquer outro registro da escrituração.
'Deve escriturar o Registro “0145” a pessoa jurídica que tenha auferido receita das atividades de serviços ou da fabricação
'de produtos, relacionados nos art. 7º e 8º da Lei nº 12.546/2011, respectivamente.
'Atenção: A partir da versão 2.02, uma vez informado o registro filho “0145”, em relação ao registro pai “0140” do
'estabelecimento matriz, fica dispensada a necessidade de escriturar o registro “0145” em relação aos demais
'estabelecimento. Desta forma, ao cadastrar o registro “0145” do estabelecimento matriz, o PVA fica habilitado para
'validar a escrituração do registro “P100” tanto do estabelecimento matriz como dos demais estabelecimentos da empresa.
'No caso de não auferir quaisquer das receitas, nas hipóteses previstas em lei, não precisa ser informado o registro “0145”,
'muito menos ser escriturado o Bloco P.
'A soma dos valores informados no campo 04 (VL_REC_ATIV) e do campo 05 (VL_REC_DEMAIS_ATIV) pode ser
'menor ou igual ao valor informado no campo 03 (VL_REC_TOT), não maior.
Dim lErro As Long
Dim sChave As String, sSQL As String
Dim dTotalRec As Double, dTotalRecContrPrev As Double

On Error GoTo Erro_Processa_Registro_0145

    '******** registro 0145 *****************
    objEFD.objSpedFiscal.sTipo = "0145"
    objEFD.objSpedFiscal.lSeqPai = 0

    'Abertura de Comando
    Call objEFD.alComando0145_AbreComandos(1)
        
    sSQL = "SELECT SUM(D.ValorRecebido), SUM(D.ValorRecContrPrev) FROM SpedDocFiscais AS D WHERE D.Registro = 'F500' AND D.NumIntSped = ? AND D.FilialEmpresa = ?"
    
    lErro = Comando_Executar(objEFD.alComando0145(1), sSQL, dTotalRec, dTotalRecContrPrev, objEFD.lNumIntSped, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 211983

    lErro = Comando_BuscarPrimeiro(objEFD.alComando0145(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211984
    
    'Deve escriturar o Registro “0145” a pessoa jurídica que tenha auferido receita das atividades de serviços ou da fabricação
    'de produtos, relacionados nos art. 7º e 8º da Lei nº 12.546/2011, respectivamente
    If dTotalRecContrPrev > 0 Then
    
        Call objEFD.objTela.SPED_Status("Processando Registro 0145", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "0145")
        Call objEFD.Reg_Campo_add(2, "COD_INC_TRIB", 1) '1 – Contribuição Previdenciária apurada no período, exclusivamente com base na Receita Bruta;
        Call objEFD.Reg_Campo_add(3, "VL_REC_TOT", SFF_Num(dTotalRec))
        Call objEFD.Reg_Campo_add(4, "VL_REC_ATIV", SFF_Num(dTotalRecContrPrev))
        Call objEFD.Reg_Campo_add(5, "VL_REC_DEMAIS_ATIV", 0)
        Call objEFD.Reg_Campo_add(6, "INFO_COMPL", "")
        Call objEFD.Finaliza_Reg
                        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Processa_Registro_0145 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0145:

   Processa_Registro_0145 = gErr

   Select Case gErr
   
        Case 211983, 211984
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211985)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0145", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_0150(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0150: TABELA DE CADASTRO DO PARTICIPANTE
'Este registro tem por objetivo relacionar e cadastrar os participantes (fornecedores e clientes pessoa jurídica ou pessoa
'física) que tenham realizado operações com a empresa, objeto de registro nos Blocos A, C, D, F ou 1.
'Em relação às operações documentadas com base em Nota Fiscal Eletrônica (Código 55), no caso da pessoa jurídica
'proceder à escrituração consolidada de suas vendas (Registro C180) e/ou de suas aquisições (Registro C190), não é
'obrigatório cadastrar e relacionar no Registro 0150 o participante cujas operações estejam exclusivamente escrituradas
'nos registros C180 e C190.
'Em relação às operações documentadas com base em Nota Fiscal Eletrônica (Código 55), no caso da pessoa jurídica
'proceder à escrituração de forma individualizada por documento fiscal (Registros C100/C170) de suas vendas e/ou de
'suas aquisições, é obrigatório cadastrar e relacionar no Registro 0150 cada participante cujas operações estejam
'escrituradas nos registros C100 e C170.
Dim lErro As Long
Dim iTransportadora As Integer
Dim lCliente As Long, iFilialCli As Integer
Dim lFornecedor As Long, iFilialForn As Integer
Dim iIndice As Integer, sSelect As String, sTipo As String, sWhere As String
Dim sIE As String, tFilialForn As typeFilialFornecedor
Dim tEndereco As typeEndereco
Dim objPais As ClassPais, sCodParticipante As String
Dim objCidades As ClassCidades, lContador As Long
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0150

    '******** registro 0150 *****************
    objEFD.objSpedFiscal.sTipo = "0150"
    objEFD.objSpedFiscal.lSeqPai = 0

    'Abertura de Comando
    Call objEFD.alComando0150_AbreComandos(3)
        
    sWhere = "SELECT * FROM SpedDocFiscais AS N WHERE N.NumIntSped = ? AND FilialEmpresa = ? AND Bloco <> '' "
    
    For iIndice = 1 To 2
    
        Select Case iIndice
            Case 1
                'sSelect = "SELECT F.CodCliente, F.CodFilial, C.RazaoSocial + '-' + F.Nome, F.CGC, F.InscricaoEstadual, F.InscricaoSuframa, E.Logradouro, E.Complemento, E.Numero, E.Bairro, E.Cidade, E.CodigoPais FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E WHERE C.Codigo = F.CodCliente AND E.Codigo = F.Endereco  AND EXISTS (" & sWhere & " AND N.Cliente = F.CodCliente AND N.FilialCli = F.CodFilial )"
                sSelect = "SELECT F.CodCliente, F.CodFilial, C.RazaoSocial, F.CGC, F.InscricaoEstadual, F.InscricaoSuframa, E.Logradouro, E.Complemento, E.Numero, E.Bairro, E.Cidade, E.CodigoPais FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E WHERE C.Codigo = F.CodCliente AND E.Codigo = F.Endereco  AND EXISTS (" & sWhere & " AND N.Cliente = F.CodCliente AND N.FilialCli = F.CodFilial )"
                sTipo = "Cliente"
            Case 2
                'sSelect = "SELECT F.CodFornecedor, F.CodFilial, C.RazaoSocial + '-' + F.Nome, F.CGC, F.InscricaoEstadual, F.InscricaoSuframa, E.Logradouro, E.Complemento, E.Numero, E.Bairro, E.Cidade, E.CodigoPais FROM Fornecedores AS C, FiliaisFornecedores AS F, Enderecos AS E WHERE C.Codigo = F.CodFornecedor AND E.Codigo = F.Endereco AND EXISTS (" & sWhere & " AND N.Fornecedor = F.CodFornecedor AND N.FilialForn = F.CodFilial )"
                sSelect = "SELECT F.CodFornecedor, F.CodFilial, C.RazaoSocial, F.CGC, F.InscricaoEstadual, F.InscricaoSuframa, E.Logradouro, E.Complemento, E.Numero, E.Bairro, E.Cidade, E.CodigoPais FROM Fornecedores AS C, FiliaisFornecedores AS F, Enderecos AS E WHERE C.Codigo = F.CodFornecedor AND E.Codigo = F.Endereco AND EXISTS (" & sWhere & " AND N.Fornecedor = F.CodFornecedor AND N.FilialForn = F.CodFilial )"
                sTipo = "Fornecedor"
        End Select
        
        sChave = "Tipo: " & sTipo
        
        tFilialForn.sNome = String(STRING_MAXIMO, 0)
        tFilialForn.sCgc = String(STRING_CGC, 0)
        tFilialForn.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        tFilialForn.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        tEndereco.sLogradouro = String(STRING_ENDERECO_LOGRADOURO, 0)
        tEndereco.sComplemento = String(STRING_ENDERECO_COMPLEMENTO, 0)
        tEndereco.sBairro = String(STRING_BAIRRO, 0)
        tEndereco.sCidade = String(STRING_CIDADE, 0)
    
        lErro = Comando_Executar(objEFD.alComando0150(iIndice), sSelect, tFilialForn.lCodFornecedor, tFilialForn.iCodFilial, tFilialForn.sNome, tFilialForn.sCgc, tFilialForn.sInscricaoEstadual, tFilialForn.sInscricaoSuframa, tEndereco.sLogradouro, tEndereco.sComplemento, tEndereco.lNumero, tEndereco.sBairro, tEndereco.sCidade, tEndereco.iCodigoPais, objEFD.lNumIntSped, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 209426
    
        lErro = Comando_BuscarPrimeiro(objEFD.alComando0150(iIndice))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209427
        
        Do While lErro <> AD_SQL_SEM_DADOS
                
            sCodParticipante = left(sTipo, 1) & Format(tFilialForn.lCodFornecedor, "0000000000") & Format(tFilialForn.iCodFilial, "00000")
     
            sChave = "Tipo: " & sTipo & " Código: " & sCodParticipante
            Call objEFD.objTela.SPED_Status("Processando Registro 0150", sChave)
    
            sIE = ""
            Call Formata_String_Numero(tFilialForn.sInscricaoEstadual, sIE)
            
            Set objPais = New ClassPais
            Set objCidades = New ClassCidades
            
            objCidades.sDescricao = tEndereco.sCidade
            
            lErro = CF("Cidade_Le_Nome", objCidades)
            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
            
            'If Len(Trim(objCidades.sCodIBGE)) = 0 Then gError 209428
        
            objPais.iCodigo = tEndereco.iCodigoPais
            
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError ERRO_SEM_MENSAGEM
            
            If objPais.iCodBacen = 0 Then gError 209429
            
            If objPais.iCodigo <> PAIS_BRASIL Then
                objCidades.sCodIBGE = "9999999"
            End If
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "0150")
            Call objEFD.Reg_Campo_add(2, "COD_PART", sCodParticipante)
            Call objEFD.Reg_Campo_add(3, "NOME", SFF_Texto(tFilialForn.sNome, 100))
            Call objEFD.Reg_Campo_add(4, "COD_PAIS", CStr(objPais.iCodBacen))
            Call objEFD.Reg_Campo_add(5, "CNPJ", IIf(Len(tFilialForn.sCgc) = 14, tFilialForn.sCgc, ""))
            Call objEFD.Reg_Campo_add(6, "CPF", IIf(Len(tFilialForn.sCgc) = 11, tFilialForn.sCgc, ""))
            Call objEFD.Reg_Campo_add(7, "IE", SFF_Texto(sIE))
            Call objEFD.Reg_Campo_add(8, "COD_MUN", objCidades.sCodIBGE)
            Call objEFD.Reg_Campo_add(9, "SUFRAMA", SFF_Texto(tFilialForn.sInscricaoSuframa, 9))
            Call objEFD.Reg_Campo_add(10, "END", SFF_Texto(tEndereco.sLogradouro, 60))
            Call objEFD.Reg_Campo_add(11, "NUM", CStr(tEndereco.lNumero))
            Call objEFD.Reg_Campo_add(12, "COMPL", SFF_Texto(tEndereco.sComplemento, 60))
            Call objEFD.Reg_Campo_add(13, "BAIRRO", SFF_Texto(tEndereco.sBairro, 60))
            Call objEFD.Finaliza_Reg
                            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lContador = lContador + 1
            
            lErro = Comando_BuscarProximo(objEFD.alComando0150(iIndice))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130
            
        Loop
        
    Next
    
    Processa_Registro_0150 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0150:

   Processa_Registro_0150 = gErr

   Select Case gErr
   
        Case 209426, 209427, 204130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", gErr)
   
        Case 209428
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO2", gErr, objCidades.sDescricao)
   
        Case 209429
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_PAIS_SEM_CODBACEN", gErr, objPais.sNome)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209430)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0150", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_0190(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0190: IDENTIFICAÇÃO DAS UNIDADES DE MEDIDA
Dim lErro As Long
Dim sUM As String, sUMDesc As String, sUMDescAux As String
Dim lContador As Long
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0190

    '******************* registro 0190 **********************************
    objEFD.objSpedFiscal.sTipo = "0190"
    objEFD.objSpedFiscal.lSeqPai = 0

    sUM = String(STRING_MAXIMO, 0)
    sUMDesc = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(objEFD.lComando0190, "SELECT Sigla, MAX(Nome) AS Nome FROM UnidadesDeMedida GROUP BY Sigla ", sUM, sUMDesc)
    If lErro <> AD_SQL_SUCESSO Then gError 209431

    lErro = Comando_BuscarPrimeiro(objEFD.lComando0190)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209432

    Do While lErro <> AD_SQL_SEM_DADOS
    
        sChave = "UM: " & CStr(sUM)
        Call objEFD.objTela.SPED_Status("Processando Registro 0190", sChave)

        If UCase(sUM) = UCase(sUMDesc) Then
            sUMDescAux = sUMDesc & "."
        Else
            sUMDescAux = sUMDesc
        End If

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "0190")
        Call objEFD.Reg_Campo_add(2, "UNID", SFF_Texto(sUM))
        Call objEFD.Reg_Campo_add(3, "DESCR", SFF_Texto(sUMDescAux))
        Call objEFD.Finaliza_Reg
            
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lContador = lContador + 1

        lErro = Comando_BuscarProximo(objEFD.lComando0190)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209433

    Loop

    Exit Function

Erro_Processa_Registro_0190:

    Processa_Registro_0190 = gErr

    Select Case gErr

        Case 209431 To 209433
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_UNIDADESDEMEDIDA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209434)

    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0190", sChave)

    Exit Function

End Function

Private Function Processa_Registro_0200(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTOS E SERVIÇOS)
'Este registro tem por objetivo informar as mercadorias, serviços, produtos ou quaisquer outros itens concernentes às
'transações representativas de receitas e/ou geradoras de créditos, objeto de escrituração nos Blocos A, C, D, F ou 1.'
Dim lErro As Long, iIndice As Integer
Dim tProduto As typeProduto, sWhere As String
Dim iNatureza As Integer
Dim lContador200 As Long, lContador205 As Long, lContador206 As Long, lContador208 As Long
Dim sChave As String, lSeq As Long

On Error GoTo Erro_Processa_Registro_0200

    Call objEFD.alComando0200_AbreComandos(4)

    '******************* registro 0200 **********************************
    objEFD.objSpedFiscal.sTipo = "0200"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    sWhere = "SELECT * FROM SpedDocFiscaisItens N, SpedDocFiscais S WHERE S.TipoDoc = N.TipoDoc AND S.NumIntDoc = N.NumIntDoc AND S.NumIntSped = N.NumIntSped AND N.NumIntSped = ? AND S.FilialEmpresa = ? AND S.Bloco <> '' "

    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
    tProduto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    tProduto.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tProduto.sGenero = String(STRING_PRODUTO_GENERO, 0)
    tProduto.sCodigo = String(STRING_PRODUTO, 0)

    'pegar os dados do produto
    lErro = Comando_Executar(objEFD.alComando0200(1), "SELECT Codigo, Descricao, CodigoBarras, SiglaUMEstoque, Natureza, IPICodigo, Genero, TipoServico, ClasseUM  FROM Produtos WHERE EXISTS (" & sWhere & " AND N.Produto = Codigo ) ", _
    tProduto.sCodigo, tProduto.sDescricao, tProduto.sCodigoBarras, tProduto.sSiglaUMEstoque, tProduto.iNatureza, tProduto.sIPICodigo, tProduto.sGenero, tProduto.iTipoServico, tProduto.iClasseUM, objEFD.lNumIntSped, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209435

    lErro = Comando_BuscarPrimeiro(objEFD.alComando0200(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209436
    
    Do While lErro = AD_SQL_SUCESSO
    
        sChave = "Produto: " & tProduto.sCodigo
        Call objEFD.objTela.SPED_Status("Processando Registro 0200", sChave)
        
        Select Case tProduto.iNatureza
            Case 1 'Materia Prima
                iNatureza = 1
            Case 2 'Produto Intermediario
                iNatureza = 6
            Case 3 'Embalagem
                iNatureza = 2
            Case 4 'Produto Acabado
                iNatureza = 4
            Case 5 'Revenda
                iNatureza = 0
            Case 6, 7 'Manutencao, Outros
                iNatureza = 99
            Case 8 'Serviços
                iNatureza = 9
            Case 9 'Em Processo
                iNatureza = 3
            Case 10 'SubProduto
                iNatureza = 5
            Case 11 'Uso e Consumo
                iNatureza = 7
            Case 12 'Ativo
                iNatureza = 8
            Case 13 'Outros Insumos
                iNatureza = 10
            Case Else
                gError 99999
        End Select
            
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "0200")
        Call objEFD.Reg_Campo_add(2, "COD_ITEM", SFF_Texto(tProduto.sCodigo))
        Call objEFD.Reg_Campo_add(3, "DESCR_ITEM", SFF_Texto(tProduto.sDescricao))
        Call objEFD.Reg_Campo_add(4, "COD_BARRA", SFF_Texto(tProduto.sCodigoBarras))
        Call objEFD.Reg_Campo_add(5, "COD_ANT_ITEM", "")
        Call objEFD.Reg_Campo_add(6, "UNID_INV", SFF_Texto(tProduto.sSiglaUMEstoque))
        Call objEFD.Reg_Campo_add(7, "TIPO_ITEM", Format(iNatureza, "00"))
        Call objEFD.Reg_Campo_add(8, "COD_NCM", SFF_Texto(tProduto.sIPICodigo, 8))
        Call objEFD.Reg_Campo_add(9, "EX_IPI", "")
        Call objEFD.Reg_Campo_add(10, "COD_GEN", IIf(tProduto.sGenero = "0", "", SFF_Texto(tProduto.sGenero)))
        Call objEFD.Reg_Campo_add(11, "COD_LST", SFF_Texto(tProduto.sISSQN))
        Call objEFD.Reg_Campo_add(12, "ALIQ_ICMS", "")
        Call objEFD.Finaliza_Reg
        
        objEFD.objSpedFiscal.sTipo = "0200"
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lContador200 = lContador200 + 1
        
        lSeq = objEFD.objSpedFiscal.lSeq

        lErro = Processa_Registro_0205(objEFD, tProduto.sCodigo, lSeq, lContador205, objEFD.alComando0200(2))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_0206(objEFD, tProduto.sCodigo, lSeq, lContador206, objEFD.alComando0200(3))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_0208(objEFD, tProduto.sCodigo, lSeq, lContador208, objEFD.alComando0200(4))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(objEFD.alComando0200(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209437
        
    Loop

    Processa_Registro_0200 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0200:

    Processa_Registro_0200 = gErr
    
    Select Case gErr

        Case 209435 To 209437
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209438)
     
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0200", sChave)
    
    Exit Function

End Function

Private Function Processa_Registro_0205(ByVal objEFD As ClassEFDPisCofinsSel, sProduto As String, lSeqPai As Long, lContador As Long, lComando As Long) As Long
'REGISTRO 0205: ALTERAÇÃO DO ITEM
'Este registro tem por objetivo informar alterações ocorridas na descrição do produto, desde que não o descaracterize ou
'haja modificação que o identifique como sendo novo produto, caso não tenha ocorrido movimentação no período da
'alteração do item, deverá ser informada no primeiro período em que houver movimentação do item.
'Deverá ser ainda informado quando ocorrer alteração na codificação do produto.
'Não podem ser informados dois ou mais registros com sobreposição de períodos.
Dim lErro As Long
Dim tProduto As typeProduto
Dim dtDataAnt As Date, dtDataAtual As Date
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0205

    '******************* registro 0205 **********************************
    objEFD.objSpedFiscal.sTipo = "0205"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "Produto: " & sProduto

    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    dtDataAnt = DateAdd("d", -1, objEFD.dtDataFim)

    'pegar os dados do produtohistorico
    lErro = Comando_Executar(lComando, "SELECT DataAtualizacao, DescProduto FROM ProdutoHistorico WHERE CodProduto = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC, NumIntDoc DESC ", _
    dtDataAtual, tProduto.sDescricao, sProduto, objEFD.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 209439

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209440

    Do While lErro = AD_SQL_SUCESSO
       
        sChave = "Produto: " & sProduto
        Call objEFD.objTela.SPED_Status("Processando Registro 0205", sChave)
            
        If objEFD.dtDataIni > dtDataAtual Then dtDataAtual = objEFD.dtDataIni
        
        'Se o mesmo produto foi alterado diversas vezes no mesmo dia só considera
        ' a última mexida
        If dtDataAnt >= dtDataAtual Then
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "0205")
            Call objEFD.Reg_Campo_add(2, "DESCR_ANT_ITEM", SFF_Texto(tProduto.sDescricao))
            Call objEFD.Reg_Campo_add(3, "DT_INI", SFF_Data(dtDataAtual)) 'A validade começa a partir de seu cadastro
            Call objEFD.Reg_Campo_add(4, "DT_FIM", SFF_Data(dtDataAnt)) 'A validade termina quando começa a de outra descrição mais recente
            Call objEFD.Reg_Campo_add(5, "COD_ANT_ITEM", "")
            Call objEFD.Finaliza_Reg
        
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        If objEFD.dtDataIni = dtDataAtual Then Exit Do
        
        dtDataAnt = DateAdd("d", -1, dtDataAtual) 'Não pode ficar com a data final igual a data inicial de outro
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209441

    Loop

    Processa_Registro_0205 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0205:

    Processa_Registro_0205 = gErr
    
    Select Case gErr

        Case 209439 To 209441
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOHISTORICO", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209442)
     
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0205", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_0206(ByVal objEFD As ClassEFDPisCofinsSel, sProduto As String, lSeqPai As Long, lContador As Long, lComando As Long) As Long
'REGISTRO 0206: CÓDIGO DE PRODUTO CONFORME TABELA ANP (COMBUSTÍVEIS)
'Este registro tem por objetivo informar o código correspondente ao produto constante na Tabela da Agência Nacional de
'Petróleo (ANP) para os produtos denominados “Combustíveis”.
'Deve ser apresentado apenas pelos contribuintes produtores, importadores e distribuidores de combustíveis.
On Error GoTo Erro_Processa_Registro_0206

    '******** registro 0206 *****************
    objEFD.objSpedFiscal.sTipo = "0206"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_0206 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0206:

   Processa_Registro_0206 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209443)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0208(ByVal objEFD As ClassEFDPisCofinsSel, sProduto As String, lSeqPai As Long, lContador As Long, lComando As Long) As Long
'REGISTRO 0208: CÓDIGO DE GRUPOS POR MARCA COMERCIAL – REFRI (BEBIDAS FRIAS).
'Este registro deve ser preenchido pela pessoa jurídica industrial ou importadora de bebidas frias (cerveja, refrigerantes,
'águas, preparações compostas não alcoólicas, etc), optante do Regime Especial de Apuração da Contribuição para o
'PIS/Pasep e da Cofins por litro de produto, conforme as alíquotas específicas por produto e marcas comerciais
'estabelecidas pelo Poder Executivo, nos termos da Lei nº 10.833, de 2003.
On Error GoTo Erro_Processa_Registro_0208

    '******** registro 0208 *****************
    objEFD.objSpedFiscal.sTipo = "0208"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_0208 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0208:

   Processa_Registro_0208 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209444)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_0400(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0400: TABELA DE NATUREZA DA OPERAÇÃO/PRESTAÇÃO
'Este registro tem por objetivo codificar os textos das diferentes naturezas da operação/prestação discriminadas nos
'documentos fiscais. Esta codificação e suas descrições são livremente criadas e mantidas pelo contribuinte.
'Este registro não se refere a CFOP. Algumas empresas utilizam outra classificação além das apresentados nos CFOP. Esta
'codificação permite informar estes agrupamentos próprios.
'Não podem ser informados dois ou mais registros com o mesmo código no campo COD_NAT
On Error GoTo Erro_Processa_Registro_0400

    '******** registro 0206 *****************
    objEFD.objSpedFiscal.sTipo = "0400"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    Processa_Registro_0400 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0400:

   Processa_Registro_0400 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209445)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_0450(ByVal objEFD As ClassEFDPisCofinsSel, ByVal iFilialEmpresa As Integer) As Long
'REGISTRO 0450: TABELA DE INFORMAÇÃO COMPLEMENTAR DO DOCUMENTO FISCAL
'Este registro tem por objetivo codificar todas as informações complementares dos documentos fiscais, exigidas pela
'legislação fiscal. Estas informações constam no campo “Dados Adicionais” dos documentos fiscais.
'Esta codificação e suas descrições são livremente criadas e mantidas pelo contribuinte e não podem ser informados dois
'ou mais registros com o mesmo conteúdo no campo COD_INF.
'Deverão constar todas as informações complementares de interesse da Administração Tributária, existentes nos
'documentos fiscais.
'Exemplo: nos casos de documentos fiscais de entradas, informar as referências a um outro documento fiscal.
Dim lErro As Long
Dim sMensagem As String, sMgsCompleta As String
Dim lNumIntDoc As Long
Dim iIndice As Integer, lCodMsg As Long
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0450

    '******************* registro 0450 **********************************
    objEFD.objSpedFiscal.sTipo = "0450"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    'Abertura de Comando
    Call objEFD.alComando0450_AbreComandos(3)
            
    'Le os documentos fiscais
    lErro = Comando_ExecutarPos(objEFD.alComando0450(1), "SELECT NumIntDoc FROM SpedDocFiscais WHERE NumIntSped = ? AND TipoDoc = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objEFD.lNumIntSped, SPED_DOC_NF, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209446

    lErro = Comando_BuscarPrimeiro(objEFD.alComando0450(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209447
    
    lCodMsg = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sMgsCompleta = ""
        
        sChave = "TipoDoc: 0 ID: " & CStr(lNumIntDoc)
        
        sMensagem = String(STRING_MAXIMO, 0)
    
        'Para cada documento monta a mensagem
        lErro = Comando_Executar(objEFD.alComando0450(2), "SELECT Mensagem FROM MensagensRegra WHERE TipoDoc = 0 AND NumIntDoc = ? ORDER BY TipoMsg, Seq", sMensagem, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 209448
    
        lErro = Comando_BuscarPrimeiro(objEFD.alComando0450(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209449
    
        Do While lErro <> AD_SQL_SEM_DADOS
                
            sMgsCompleta = sMgsCompleta & Replace(sMensagem, "|", "")
    
            lErro = Comando_BuscarProximo(objEFD.alComando0450(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209450
    
        Loop
        
        If Len(Trim(sMgsCompleta)) > 0 Then
        
            lCodMsg = lCodMsg + 1
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "0450")
            Call objEFD.Reg_Campo_add(2, "COD_INF", Format(lCodMsg, "000000"))
            Call objEFD.Reg_Campo_add(3, "TXT", SFF_Texto(sMgsCompleta, 255))
            Call objEFD.Finaliza_Reg
    
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_ExecutarPos(objEFD.alComando0450(3), "UPDATE SpedDocFiscais SET CodMensagem = ?", objEFD.alComando0450(1), lCodMsg)
            If lErro <> AD_SQL_SUCESSO Then gError 209451
            
        End If
    
        lErro = Comando_BuscarProximo(objEFD.alComando0450(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209452
    
    Loop
    
    Exit Function

Erro_Processa_Registro_0450:

    Processa_Registro_0450 = gErr

    Select Case gErr

        Case 209446, 209447, 209452
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case 209448 To 209450
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENSAGENSREGRAS", gErr)

        Case 209451
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209453)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0450", sChave)
    
    Exit Function

End Function

Private Function Processa_Registro_0500(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'Este registro tem o objetivo de identificar as contas contábeis utilizadas pelo contribuinte informante em sua Escrituração
'Contábil, relacionadas às operações representativas de receitas, tributadas ou não, e dos créditos apurados.
'Não podem ser informados dois ou mais registros com a mesma combinação de conteúdo nos campos DT_ALT e
'COD_CTA.
Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lContador As Long, tConta As typePlanoConta
Dim sTipo As String, sConta As String, dtDataAtualizacao As Date
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0500

    '******************* registro 0500 **********************************
    objEFD.objSpedFiscal.sTipo = "0500"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
        
    tConta.sConta = String(STRING_CONTA, 0)
    tConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Conta, NivelConta, DescConta, TipoConta, NaturezaSped FROM PlanoConta ORDER BY Conta", _
                                            tConta.sConta, tConta.iNivelConta, tConta.sDescConta, tConta.iTipoConta, tConta.iNaturezaSped)
    If lErro <> AD_SQL_SUCESSO Then gError 209454

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209455

    Do While lErro <> AD_SQL_SEM_DADOS
    
        lContador = lContador + 1
        
        sChave = "Conta: " & tConta.sConta
        Call objEFD.objTela.SPED_Status("Processando Registro 0500", sChave)
    
        If tConta.iNaturezaSped = 0 Then gError 209456 'Tem que estar preenchido
        
        Call Mascara_RetornaContaEnxuta(tConta.sConta, sConta)
        
        If tConta.iTipoConta = CONTA_SINTETICA Then
            sTipo = "S"
        Else
            sTipo = "A"
        End If
        
        lErro = Comando_Executar(alComando(2), "SELECT MAX(DataAtualizacao) FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ?", _
                                                dtDataAtualizacao, tConta.sConta, objEFD.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 209457
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209458
        
        'Se não encontrar é porque foi cadastrada após a data final do SPED
        If dtDataAtualizacao <> DATA_NULA Then
            'Por enquanto não vai pegar a conta ref porque pelo sistema uma mesma conta pode estar associada a mais de
            'uma conta referencial dependendo do centro de custo
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "0500")
            Call objEFD.Reg_Campo_add(2, "DT_ALT", SFF_Data(dtDataAtualizacao))
            Call objEFD.Reg_Campo_add(3, "COD_NAT_CC", Format(tConta.iNaturezaSped, "00"))
            Call objEFD.Reg_Campo_add(4, "IND_CTA", sTipo)
            Call objEFD.Reg_Campo_add(5, "NÍVEL", CStr(tConta.iNivelConta))
            Call objEFD.Reg_Campo_add(6, "COD_CTA", sConta)
            Call objEFD.Reg_Campo_add(7, "NOME_CTA", SFF_Texto(tConta.sDescConta, 60))
            Call objEFD.Reg_Campo_add(8, "COD_CTA_REF", "")
            Call objEFD.Reg_Campo_add(9, "CNPJ_EST", "")
            Call objEFD.Finaliza_Reg
            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209459
    Loop

    Exit Function

Erro_Processa_Registro_0500:

    Processa_Registro_0500 = gErr

    Select Case gErr

        Case 209454, 209455, 209459
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)

        Case 209456
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEM_NATUREZA_SPED", gErr, tConta.sConta)

        Case 209457, 209458
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209459)

    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0500", sChave)
    
    Exit Function

End Function

Private Function Processa_Registro_0600(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0600: CENTRO DE CUSTOS
'Este registro tem o objetivo de identificar os centros de custos referenciados nos registros de operações e documentos
'escriturados na EFD-PIS/Cofins.
'Não podem ser informados dois ou mais registros com a mesma combinação de conteúdo nos campos DT_ALT e
'COD_CCUS.
Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long
Dim lContador As Long, tCcl As typeCcl
Dim sCcl As String, dtDataAtualizacao As Date, dtDataMin As Date
Dim sChave As String

On Error GoTo Erro_Processa_Registro_0600

    '******************* registro 0600 **********************************
    objEFD.objSpedFiscal.sTipo = "0600"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
     lErro = Comando_Executar(alComando(3), "SELECT MIN(DataAtualizacao) FROM CclHistorico WHERE DataAtualizacao <> ?", _
                                            dtDataMin, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 209497

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209498
    
    If dtDataMin = DATA_NULA Then dtDataMin = objEFD.dtDataIni
    
    tCcl.sCcl = String(STRING_CCL, 0)
    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Ccl, DescCcl FROM Ccl ORDER BY Ccl", _
                                            tCcl.sCcl, tCcl.sDescCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 209460

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209461

    Do While lErro <> AD_SQL_SEM_DADOS
           
        sChave = "Ccl: " & tCcl.sCcl
        Call objEFD.objTela.SPED_Status("Processando Registro 0600", sChave)
           
        lContador = lContador + 1
          
        Call Mascara_RetornaCclEnxuta(tCcl.sCcl, sCcl)
        
        lErro = Comando_Executar(alComando(2), "SELECT MAX(DataAtualizacao) FROM CclHistorico WHERE CodigoCcl = ? AND DataAtualizacao <= ?", _
                                                dtDataAtualizacao, tCcl.sCcl, objEFD.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 209462
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209463
        
        'If dtDataAtualizacao = DATA_NULA Then dtDataAtualizacao = dtDataMin

        'Se não encontrar é porque foi cadastrada após a data final do SPED
        If dtDataAtualizacao <> DATA_NULA Then
        
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "0600")
            Call objEFD.Reg_Campo_add(2, "DT_ALT", SFF_Data(dtDataAtualizacao))
            Call objEFD.Reg_Campo_add(3, "COD_CCUS", sCcl)
            Call objEFD.Reg_Campo_add(4, "CCUS", SFF_Texto(tCcl.sDescCcl, 60))
            Call objEFD.Finaliza_Reg
            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209464
    Loop

    Exit Function

Erro_Processa_Registro_0600:

    Processa_Registro_0600 = gErr

    Select Case gErr

        Case 209460, 209461, 209464, 209497, 209498
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", gErr)

        Case 209462, 209463
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209465)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_0600", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_0990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_0990

    '******************* registro 0990 **********************************
    objEFD.objSpedFiscal.sTipo = "0990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1 'ele mesmo

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "0" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "0990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_0", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_0990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_0990:

   Processa_Registro_0990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209466)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_A001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBlocoA As Boolean) As Long
'REGISTRO A001: ABERTURA DO BLOCO A
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_A001

    '******** registro A001 *****************
    
    'Se  for a versão inferior a 3 ou a filial não seja lucro presumido ou não cumulativo ou sem regime de caixa, lê doc a doc
    If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA Then
        
        lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE NumIntSped = ? AND Bloco = ?", lContador, objEFD.lNumIntSped, "A")
        If lErro <> AD_SQL_SUCESSO Then gError 209467
    
        lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209468
        
    End If
        
    objEFD.objSpedFiscal.sTipo = "A001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBlocoA = True
    Else
        iTemReg = 1
        bAbreBlocoA = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "A001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_A001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A001:

   Processa_Registro_A001 = gErr

   Select Case gErr
   
        Case 209467, 209468
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209469)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_A010(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO A010: IDENTIFICAÇÃO DO ESTABELECIMENTO
'Este registro tem o objetivo de identificar o estabelecimento da pessoa jurídica a que se referem as operações e
'documentos fiscais informados neste bloco. Só devem ser escriturados no Registro A010 os estabelecimentos que
'efetivamente tenham realizado operações de prestação ou de contratação de serviços, mediante emissão de documento
'fiscal, que devam ser escrituradas no Bloco A.
'O estabelecimento que não realizou operações passíveis de registro nesse bloco, no período da escrituração, não deve ser
'identificado no Registro A010.
'Para cada estabelecimento cadastrado em “A010”, deve ser informado nos registros de nível inferior (Registros Filho) as
'operações próprias de prestação ou de contratação de serviços, mediante emissão de documento fiscal, no mercado
'interno ou externo.
Dim lErro As Long, iIndice As Integer
Dim lContA010 As Long, lContA100 As Long, lContA110 As Long
Dim lContA111 As Long, lContA120 As Long, lContA170 As Long
Dim objDoc As ClassSpedDocFiscais
Dim objDocItens As ClassSpedDocFiscaisItens
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 2) As Long
Dim sCGCAnt As String, sChave As String

On Error GoTo Erro_Processa_Registro_A010

    '******************* registro A010 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    sCGCAnt = "X"  'Força ser diferente
    
    lErro = SpedFiscal_Docs_Prepara(alComando(0), objEFD.lNumIntSped, "A", tDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209470
   
    Do While lErro = AD_SQL_SUCESSO
        
        Set objDoc = New ClassSpedDocFiscais
        
        lErro = SpedFiscal_Doc_Copia(objDoc, tDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = SpedFiscal_DocItens_Le(alComando(1), objDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Mudou o CPF\CNPJ grava o A010
        If sCGCAnt <> objDoc.sCgc Then
            
            sChave = "CNPJ\CPF: " & objDoc.sCgc
            Call objEFD.objTela.SPED_Status("Processando Registro A010", sChave)
            
            objEFD.objSpedFiscal.sTipo = "A010"
            objEFD.objSpedFiscal.lSeqPai = 0
            
            lContA010 = lContA010 + 1
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "A010")
            Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(objDoc.sCgc, 14))
            Call objEFD.Finaliza_Reg
            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sCGCAnt = objDoc.sCgc
        End If
        
        lContA100 = lContA100 + 1
        lContA170 = lContA170 + objDoc.colItens.Count
        
        lErro = Processa_Registro_A100(objEFD, objDoc, objEFD.objSpedFiscal.lSeq, lContA110, lContA111, lContA120)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209471
    
    Loop

    Processa_Registro_A010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A010:

   Processa_Registro_A010 = gErr

   Select Case gErr
   
        Case 209470, 209471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209472)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A010", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_A100(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContA110 As Long, lContA111 As Long, lContA120 As Long) As Long
'REGISTRO A100: DOCUMENTO - NOTA FISCAL DE SERVIÇO
'Deve ser gerado um Registro A100 para cada documento fiscal a ser relacionado na escrituração, referente à prestação ou
'à contratação de serviços, que envolvam a emissão de documentos fiscais estabelecidos pelos Municípios, eletrônicos ou
'em papel.
'Para cada registro A100, obrigatoriamente deve ser apresentado, pelo menos, um registro A170.
'Não podem ser informados, para um mesmo documento fiscal, dois ou mais registros com a mesma combinação de
'valores dos campos formadores da chave do registro. A chave deste registro é:
'• para documentos com campo IND_EMIT igual a “1” (um) – emissão por terceiros: campo IND_OPER, campo
'IND_EMIT, campo COD_PART, campo COD_SIT, campo SER e campo NUM_DOC;
'• para documentos com campo (IND_EMIT igual “0” (zero) – emissão própria: campo IND_OPER, campo IND_EMIT,
'campo COD_SIT, campo SER e campo NUM_DOC.
Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String

On Error GoTo Erro_Processa_Registro_A100

    '******** registro A100 *****************
    objEFD.objSpedFiscal.sTipo = "A100"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro A100", sChave)
    
    If objDoc.lCliente > 0 Then
        sCodParticipante = "C" & Format(objDoc.lCliente, "0000000000") & Format(objDoc.iFilialCli, "00000")
    Else
        sCodParticipante = "F" & Format(objDoc.lFornecedor, "0000000000") & Format(objDoc.iFilialForn, "00000")
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "A100")
    Call objEFD.Reg_Campo_add(2, "IND_OPER", objDoc.sOperacao)
    Call objEFD.Reg_Campo_add(3, "IND_EMIT", objDoc.sEmitente)
    If objDoc.iSituacao <> 2 Then 'Se não é cancelado
        Call objEFD.Reg_Campo_add(4, "COD_PART", sCodParticipante)
    Else
        Call objEFD.Reg_Campo_add(4, "COD_PART", "")
    End If
    Call objEFD.Reg_Campo_add(5, "COD_SIT", Format(objDoc.iSituacao, "00"))
    Call objEFD.Reg_Campo_add(6, "SER", objDoc.sSerie)
    Call objEFD.Reg_Campo_add(7, "SUB", "")
    Call objEFD.Reg_Campo_add(8, "NUM_DOC", CStr(objDoc.lNumDocumento))
    Call objEFD.Reg_Campo_add(9, "CHV_NFSE", objDoc.sChaveEletronica)
    
    If objDoc.iSituacao <> 2 Then 'Se não é cancelado
        Call objEFD.Reg_Campo_add(10, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
        Call objEFD.Reg_Campo_add(11, "DT_EXE_SERV", SFF_Data(objDoc.dtDataES))
        Call objEFD.Reg_Campo_add(12, "VL_DOC", SFF_Num(objDoc.dValorTotal))
        Call objEFD.Reg_Campo_add(13, "IND_PGTO", objDoc.sPagamento)
        Call objEFD.Reg_Campo_add(14, "VL_DESC", SFF_Num(objDoc.dValorDesconto))
        Call objEFD.Reg_Campo_add(15, "VL_BC_PIS", SFF_Num(objDoc.dValorBasePIS))
        Call objEFD.Reg_Campo_add(16, "VL_PIS", SFF_Num(objDoc.dValorPIS))
        Call objEFD.Reg_Campo_add(17, "VL_BC_COFINS", SFF_Num(objDoc.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(18, "VL_COFINS", SFF_Num(objDoc.dValorCofins))
        Call objEFD.Reg_Campo_add(19, "VL_PIS_RET", SFF_Num(objDoc.dValorPisRetido))
        Call objEFD.Reg_Campo_add(20, "VL_COFINS_RET", SFF_Num(objDoc.dValorCofinsRetido))
        Call objEFD.Reg_Campo_add(21, "VL_ISS", SFF_Num(objDoc.dValorISS))
    Else
        Call objEFD.Reg_Campo_add(10, "DT_DOC", "")
        Call objEFD.Reg_Campo_add(11, "DT_EXE_SERV", "")
        Call objEFD.Reg_Campo_add(12, "VL_DOC", "")
        Call objEFD.Reg_Campo_add(13, "IND_PGTO", "")
        Call objEFD.Reg_Campo_add(14, "VL_DESC", "")
        Call objEFD.Reg_Campo_add(15, "VL_BC_PIS", "")
        Call objEFD.Reg_Campo_add(16, "VL_PIS", "")
        Call objEFD.Reg_Campo_add(17, "VL_BC_COFINS", "")
        Call objEFD.Reg_Campo_add(18, "VL_COFINS", "")
        Call objEFD.Reg_Campo_add(19, "VL_PIS_RET", "")
        Call objEFD.Reg_Campo_add(20, "VL_COFINS_RET", "")
        Call objEFD.Reg_Campo_add(21, "VL_ISS", "")
    End If
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objDoc.iSituacao <> 2 Then 'Se não é cancelado
        lErro = Processa_Registro_A110(objEFD, objDoc, lSeq, lContA110)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Processa_Registro_A111(objEFD, objDoc, lSeq, lContA111)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Processa_Registro_A120(objEFD, objDoc, lSeq, lContA120)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Processa_Registro_A170(objEFD, objDoc, lSeq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Processa_Registro_A100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A100:

   Processa_Registro_A100 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209473)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A100", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_A110(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContA110 As Long) As Long
'REGISTRO A110: COMPLEMENTO DO DOCUMENTO - INFORMAÇÃO COMPLEMENTAR DA NF
'Este registro tem por objetivo identificar os dados contidos no campo Informações Complementares da Nota Fiscal, que
'sejam de interesse do Fisco ou conforme disponha a legislação, e que estejam explicitamente citadas no documento
'Fiscal, tais como: forma de pagamento, local da prestação/execução do serviço, operação realizada com suspensão das
'contribuições sociais, etc.
'Não podem ser informados para um mesmo documento fiscal, dois ou mais registros com o mesmo conteúdo no campo
'COD_INF.
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_A110

    '******** registro A110 *****************
    objEFD.objSpedFiscal.sTipo = "A110"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    If objDoc.lCodMensagem > 0 Then
            
        lContA110 = lContA110 + 1
            
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "A110")
        Call objEFD.Reg_Campo_add(2, "COD_INF", Format(objDoc.lCodMensagem, "000000"))
        Call objEFD.Reg_Campo_add(3, "TXT_COMPL", "")
        Call objEFD.Finaliza_Reg
            
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Processa_Registro_A110 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A110:

   Processa_Registro_A110 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209474)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_A111(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContA111 As Long) As Long
'REGISTRO A111: PROCESSO REFERENCIADO

On Error GoTo Erro_Processa_Registro_A111

    '******** registro A111 *****************
    objEFD.objSpedFiscal.sTipo = "A111"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_A111 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A111:

   Processa_Registro_A111 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209475)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_A120(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContA120 As Long) As Long
'REGISTRO A120: INFORMAÇÃO COMPLEMENTAR - OPERAÇÕES DE IMPORTAÇÃO
'Este registro tem por objetivo informar detalhes das operações de importação de serviços com direito a crédito, referentes
'a documento fiscal escriturado em A100 e que no registro filho A170 conste CST_PIS ou CST_COFINS gerador de
'crédito (CST 50 a 56), bem como conste ser o registro A170 originário de operação de importação (campo
'IND_ORIG_CRED = 1#

On Error GoTo Erro_Processa_Registro_A120

    '******** registro A120*****************
    objEFD.objSpedFiscal.sTipo = "A120"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_A120 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A120:

   Processa_Registro_A120 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209476)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_A170(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long) As Long
'REGISTRO A170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO
'Registro obrigatório para discriminar os itens da nota fiscal de serviço emitida pela pessoa jurídica ou por terceiros.
'Não podem ser informados para um mesmo documento fiscal, dois ou mais registros com o mesmo conteúdo no campo
'NUM_ITEM.
Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim sChave As String

On Error GoTo Erro_Processa_Registro_A170

    '******** registro A170 *****************
    objEFD.objSpedFiscal.sTipo = "A170"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    For Each objItem In objDoc.colItens
    
        sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Item: " & CStr(objItem.iItem)
        Call objEFD.objTela.SPED_Status("Processando Registro A170", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "A170")
        Call objEFD.Reg_Campo_add(2, "NUM_ITEM", Format(objItem.iItem, "000"))
        Call objEFD.Reg_Campo_add(3, "COD_ITEM", SFF_Texto(objItem.sProduto))
        Call objEFD.Reg_Campo_add(4, "DESCR_COMPL", SFF_Texto(objItem.sDescricao))
        Call objEFD.Reg_Campo_add(5, "VL_ITEM", SFF_Num(objItem.dPrecoTotal))
        Call objEFD.Reg_Campo_add(6, "VL_DESC", SFF_Num(objItem.dValorDesconto))
        Call objEFD.Reg_Campo_add(7, "NAT_BC_CRED", SFF_Texto(objItem.sNatBCCred))
        Call objEFD.Reg_Campo_add(8, "IND_ORIG_CRED", CStr(IIf(left(objItem.sCFOP, 1) = "3", 1, 0)))
        Call objEFD.Reg_Campo_add(9, "CST_PIS", objItem.sPISCST)
        Call objEFD.Reg_Campo_add(10, "VL_BC_PIS", SFF_Num(objItem.dValorBasePIS))
        Call objEFD.Reg_Campo_add(11, "ALIQ_PIS", SFF_Perc(objItem.dPISAliquota))
        Call objEFD.Reg_Campo_add(12, "VL_PIS", SFF_Num(objItem.dValorPIS))
        Call objEFD.Reg_Campo_add(13, "CST_COFINS", objItem.sCofinsCST)
        Call objEFD.Reg_Campo_add(14, "VL_BC_COFINS", SFF_Num(objItem.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(15, "ALIQ_COFINS", SFF_Perc(objItem.dCOFINSAliquota))
        Call objEFD.Reg_Campo_add(16, "VL_COFINS", SFF_Num(objItem.dValorCofins))
        Call objEFD.Reg_Campo_add(17, "COD_CTA", "")
        Call objEFD.Reg_Campo_add(18, "COD_CCUS", "")
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
    
    Processa_Registro_A170 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A170:

   Processa_Registro_A170 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209477)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A170", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_A990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO A990: ENCERRAMENTO DO BLOCO A
'Este registro destina-se a identificar o encerramento do bloco A e informar a quantidade de linhas (registros)
'existentes no bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_A990

    '******************* registro A990 **********************************
    objEFD.objSpedFiscal.sTipo = "A990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1 ' ele mesmo
        
    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "A" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "A990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_A", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_A990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_A990:

   Processa_Registro_A990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209478)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_C001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBlocoC As Boolean) As Long
'REGISTRO C001: ABERTURA DO BLOCO C
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_C001

    '******** registro C001 *****************

    'Se  for a versão inferior a 3 ou a filial não seja lucro presumido ou não cumulativo ou sem regime de caixa, lê doc a doc
    If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA Then
    
        lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE NumIntSped = ? AND Bloco = ?", lContador, objEFD.lNumIntSped, "C")
        If lErro <> AD_SQL_SUCESSO Then gError 209479
    
        lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209480
        
    End If
        
    objEFD.objSpedFiscal.sTipo = "C001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBlocoC = True
    Else
        iTemReg = 1
        bAbreBlocoC = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "C001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_C001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C001:

   Processa_Registro_C001 = gErr

   Select Case gErr
   
        Case 209479, 209480
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209481)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_C010(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO C010: IDENTIFICAÇÃO DO ESTABELECIMENTO
'Este registro tem o objetivo de identificar o estabelecimento da pessoa jurídica a que se referem as operações e
'documentos fiscais informados neste bloco. Só devem ser escriturados no Registro C010 os estabelecimentos que
'efetivamente tenham realizado aquisição, venda ou devolução de mercadorias, bens e produtos, mediante emissão de
'documento fiscal definido pela legislação do ICMS e do IPI, que devam ser escrituradas no Bloco C.
'O estabelecimento que não realizou operações passíveis de registro nesse bloco, no período da escrituração, não deve ser
'identificado no Registro C010.
'Para cada estabelecimento cadastrado em “C010”, deve ser informado nos registros de nível inferior (Registros Filho) as
'operações próprias de prestação ou de contratação de serviços, mediante emissão de documento fiscal, no mercado
'interno ou externo.
Dim lErro As Long, iIndice As Integer
Dim lContC010 As Long, lContC100 As Long, lContC110 As Long
Dim lContC111 As Long, lContC120 As Long, lContC170 As Long
Dim lContC500 As Long, lContC501 As Long, lContC505 As Long, lContC509 As Long
Dim objDoc As ClassSpedDocFiscais
Dim objDocItens As ClassSpedDocFiscaisItens
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 3) As Long
Dim sCGCAnt As String, sChave As String

On Error GoTo Erro_Processa_Registro_C010

    '******************* registro C010 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    sCGCAnt = "X"  'Força ser diferente
    
    lErro = SpedFiscal_Docs_Prepara(alComando(0), objEFD.lNumIntSped, "C", tDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209482
   
    Do While lErro = AD_SQL_SUCESSO
        
        Set objDoc = New ClassSpedDocFiscais
        
        lErro = SpedFiscal_Doc_Copia(objDoc, tDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = SpedFiscal_DocItens_Le(alComando(1), objDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Mudou o CPF\CNPJ grava o C010
        If sCGCAnt <> objDoc.sCgc Then
            
            sChave = "CNPJ\CPF: " & objDoc.sCgc
            Call objEFD.objTela.SPED_Status("Processando Registro C010", sChave)
            
            objEFD.objSpedFiscal.sTipo = "C010"
            objEFD.objSpedFiscal.lSeqPai = 0
            
            lContC010 = lContC010 + 1
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "C010")
            Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(objDoc.sCgc, 14))
            Call objEFD.Reg_Campo_add(3, "IND_ESCRI", "2")
            Call objEFD.Finaliza_Reg
            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sCGCAnt = objDoc.sCgc
        End If
                
        Select Case objDoc.sRegistro
        
            Case "C100" 'visão documental
            
                lContC100 = lContC100 + 1
            
                lContC170 = lContC170 + objDoc.colItens.Count
                
                lErro = Processa_Registro_C100(objEFD, objDoc, objEFD.objSpedFiscal.lSeq, lContC110, lContC111, lContC120, alComando(3))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case "C180", "C190" 'visão consolidada (Está tratando somente a visão documental)
            
            Case "C400", "C490" 'Cupom Fiscal, falta fazer
            
            Case "C500" 'Aquisição de água\luz e gás
            
                lContC500 = lContC500 + 1
            
                lErro = Processa_Registro_C500(objEFD, objDoc, objEFD.objSpedFiscal.lSeq, lContC501, lContC505, lContC509)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case "C600" 'Fornecimento de água\luz e gás (Não trata)
            
        End Select
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209483
    
    Loop

    Processa_Registro_C010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C010:

   Processa_Registro_C010 = gErr

   Select Case gErr
   
        Case 209482, 209483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209484)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C010", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C100(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC110 As Long, lContC111 As Long, lContC120 As Long, ByVal lComando As Long) As Long
'REGISTRO C100: DOCUMENTO - NOTA FISCAL (CÓDIGO 01), NOTA FISCAL AVULSA
'(CÓDIGO 1B), NOTA FISCAL DE PRODUTOR (CÓDIGO 04) e NF-e (CÓDIGO 55)
'Registro com estrutura, campos e conteúdo definidos e constantes no Leiaute da Escrituração Fiscal Digital – EFD (ICMS
'e IPI), instituído pelo Ato COTEPE/ICMS nº 9, de 12 de abril de 2008, disponível no portal de serviços (SPED) da
'página da Secretaria da Receita Federal do Brasil na Internet, no endereço <www.receita.fazenda.gov.br/sped>.
'Este registro deve ser gerado para cada documento fiscal código 01, 1B, 04 e 55, registrando a entrada ou saída de
'produtos ou outras situações que envolvam a emissão dos documentos fiscais mencionados, representativos de receitas
'auferidas, tributadas ou não pelo PIS/Pasep ou pela Cofins, bem como de operações de aquisições e/ou devoluções com
'direito a crédito da não cumulatividade.
'Não devem ser informados documentos fiscais que não se refiram a operações geradoras de receitas ou de créditos de
'PIS/Pasep e de Cofins.
'Para cada registro C100, obrigatoriamente deve ser apresentado, pelo menos, um registro C170, exceto em relação aos
'documentos fiscais referentes à nota fiscal cancelada (código “02”), Nota Fiscal Eletrônica (NF-e) denegada (código
'“04”) ou numeração inutilizada (código “05”), os quais não devem ser escriturados os registros filhos de C100.
'Não podem ser informados, para um mesmo documento fiscal, dois ou mais registros com a mesma combinação de
'valores dos campos formadores da chave do registro. A chave deste registro é:
'• para documentos com campo “IND_EMIT” igual a “1” (um) – emissão por terceiros: campo IND_OPER, campo
'IND_EMIT, campo COD_PART, campo COD_MOD, campo COD_SIT, campo SER e campo NUM_DOC;
'• para documentos com campo “IND_EMIT” igual “0” (zero) – emissão própria: campo IND_OPER, campo IND_EMIT,
'campo COD_MOD, campo COD_SIT, campo SER e campo NUM_DOC.
Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String

On Error GoTo Erro_Processa_Registro_C100

    '******** registro C100 *****************
    objEFD.objSpedFiscal.sTipo = "C100"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro C100", sChave)
    
    If objDoc.sModelo = "65" Then
        sCodParticipante = ""
    Else
        If objDoc.lCliente > 0 Then
            sCodParticipante = "C" & Format(objDoc.lCliente, "0000000000") & Format(objDoc.iFilialCli, "00000")
        Else
            sCodParticipante = "F" & Format(objDoc.lFornecedor, "0000000000") & Format(objDoc.iFilialForn, "00000")
        End If
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "C100")
    Call objEFD.Reg_Campo_add(2, "IND_OPER", objDoc.sOperacao)
    Call objEFD.Reg_Campo_add(3, "IND_EMIT", objDoc.sEmitente)
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
            Call objEFD.Reg_Campo_add(4, "COD_PART", "")
        
        Case Else
            Call objEFD.Reg_Campo_add(4, "COD_PART", sCodParticipante)
            
    End Select
    
    Call objEFD.Reg_Campo_add(5, "COD_MOD", objDoc.sModelo)
    Call objEFD.Reg_Campo_add(6, "COD_SIT", Format(objDoc.iSituacao, "00"))
    Call objEFD.Reg_Campo_add(7, "SER", objDoc.sSerie)
    Call objEFD.Reg_Campo_add(8, "NUM_DOC", CStr(objDoc.lNumDocumento))
    Call objEFD.Reg_Campo_add(9, "CHV_NFE", objDoc.sChaveEletronica)
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
            Call objEFD.Reg_Campo_add(10, "DT_DOC", "")
            Call objEFD.Reg_Campo_add(11, "DT_E_S", "")
            Call objEFD.Reg_Campo_add(12, "VL_DOC", "")
            Call objEFD.Reg_Campo_add(13, "IND_PGTO", "")
            Call objEFD.Reg_Campo_add(14, "VL_DESC", "")
            Call objEFD.Reg_Campo_add(15, "VL_ABAT_NT", "")
            Call objEFD.Reg_Campo_add(16, "VL_MERC", "")
            Call objEFD.Reg_Campo_add(17, "IND_FRT", "")
            Call objEFD.Reg_Campo_add(18, "VL_FRT", "")
            Call objEFD.Reg_Campo_add(19, "VL_SEG", "")
            Call objEFD.Reg_Campo_add(20, "VL_OUT_DA", "")
            Call objEFD.Reg_Campo_add(21, "VL_BC_ICMS", "")
            Call objEFD.Reg_Campo_add(22, "VL_ICMS", "")
            Call objEFD.Reg_Campo_add(23, "VL_BC_ICMS_ST", "")
            Call objEFD.Reg_Campo_add(24, "VL_ICMS_ST", "")
            Call objEFD.Reg_Campo_add(25, "VL_IPI", "")
            Call objEFD.Reg_Campo_add(26, "VL_PIS", "")
            Call objEFD.Reg_Campo_add(27, "VL_COFINS", "")
            Call objEFD.Reg_Campo_add(28, "VL_PIS_ST", "")
            Call objEFD.Reg_Campo_add(29, "VL_COFINS_ST", "")
        
        Case Else
    
            Call objEFD.Reg_Campo_add(10, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
            Call objEFD.Reg_Campo_add(11, "DT_E_S", SFF_Data(objDoc.dtDataES))
            Call objEFD.Reg_Campo_add(12, "VL_DOC", SFF_Num(objDoc.dValorTotal))
            Call objEFD.Reg_Campo_add(13, "IND_PGTO", objDoc.sPagamento)
            Call objEFD.Reg_Campo_add(14, "VL_DESC", SFF_Num(objDoc.dValorDesconto))
            Call objEFD.Reg_Campo_add(15, "VL_ABAT_NT", SFF_Num(objDoc.dValorNaoTributado))
            Call objEFD.Reg_Campo_add(16, "VL_MERC", SFF_Num(objDoc.dValorMercadoria))
            Call objEFD.Reg_Campo_add(17, "IND_FRT", objDoc.sFrete)
            Call objEFD.Reg_Campo_add(18, "VL_FRT", SFF_Num(objDoc.dValorFrete))
            Call objEFD.Reg_Campo_add(19, "VL_SEG", SFF_Num(objDoc.dValorSeguro))
            Call objEFD.Reg_Campo_add(20, "VL_OUT_DA", SFF_Num(objDoc.dValorDespesas))
            Call objEFD.Reg_Campo_add(21, "VL_BC_ICMS", SFF_Num(objDoc.dValorBaseICMS))
            Call objEFD.Reg_Campo_add(22, "VL_ICMS", SFF_Num(objDoc.dValorICMS))
            
            'O manual fala que não precisa ser informado esses campos
            'mas como o validador do EFD reclama quando não vão
            'o código foi comentado
            'If objDoc.sModelo = "65" Then
                'Call objEFD.Reg_Campo_add(23, "VL_BC_ICMS_ST", "")
                'Call objEFD.Reg_Campo_add(24, "VL_ICMS_ST", "")
                'Call objEFD.Reg_Campo_add(25, "VL_IPI", "")
                'Call objEFD.Reg_Campo_add(26, "VL_PIS", "")
                'Call objEFD.Reg_Campo_add(27, "VL_COFINS", "")
                'Call objEFD.Reg_Campo_add(28, "VL_PIS_ST", "")
                'Call objEFD.Reg_Campo_add(29, "VL_COFINS_ST", "")
            'Else
                Call objEFD.Reg_Campo_add(23, "VL_BC_ICMS_ST", SFF_Num(objDoc.dValorICMSST))
                Call objEFD.Reg_Campo_add(24, "VL_ICMS_ST", SFF_Num(objDoc.dValorICMSST))
                Call objEFD.Reg_Campo_add(25, "VL_IPI", SFF_Num(objDoc.dValorIPI))
                Call objEFD.Reg_Campo_add(26, "VL_PIS", SFF_Num(objDoc.dValorPIS))
                Call objEFD.Reg_Campo_add(27, "VL_COFINS", SFF_Num(objDoc.dValorCofins))
                Call objEFD.Reg_Campo_add(28, "VL_PIS_ST", SFF_Num(objDoc.dValorPisST))
                Call objEFD.Reg_Campo_add(29, "VL_COFINS_ST", SFF_Num(objDoc.dValorCofinsST))
            'End If
    
    End Select
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
        Case Else
    
            If objDoc.sModelo = "65" Then
                lErro = Processa_Registro_C175(objEFD, objDoc, lSeq)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Else
                lErro = Processa_Registro_C110(objEFD, objDoc, lSeq, lContC110)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = Processa_Registro_C111(objEFD, objDoc, lSeq, lContC111)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = Processa_Registro_C120(objEFD, objDoc, lSeq, lContC120, lComando)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                lErro = Processa_Registro_C170(objEFD, objDoc, lSeq)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
    End Select
    
    Processa_Registro_C100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C100:

   Processa_Registro_C100 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209485)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A100", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C110(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC110 As Long) As Long
'REGISTRO C110: COMPLEMENTO DO DOCUMENTO - INFORMAÇÃO COMPLEMENTAR
'DA NOTA FISCAL (CÓDIGOS 01, 1B, 04 e 55)
'Este registro tem por objetivo identificar os dados contidos no campo Informações Complementares da Nota Fiscal, que
'sejam de interesse do Fisco ou conforme disponha a legislação, e que estejam explicitamente citadas no documento
'Fiscal, tais como: forma de pagamento, local da prestação/execução do serviço, operação realizada com suspensão das
'contribuições sociais, etc.
'Não podem ser informados para um mesmo documento fiscal, dois ou mais registros com o mesmo conteúdo no campo
'COD_INF.
'Registro com estrutura, campos e conteúdo definidos e constantes no Leiaute da Escrituração Fiscal Digital – EFD
'(ICMS e IPI), instituído pelo Ato COTEPE/ICMS nº 9, de 12 de abril de 2008, disponível no portal de serviços (SPED)
'da página da Secretaria da Receita Federal do Brasil na Internet, no endereço <www.receita.fazenda.gov.br/sped>.
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_C110

    '******** registro C110 *****************
    objEFD.objSpedFiscal.sTipo = "C110"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    If objDoc.lCodMensagem > 0 Then
        
        lContC110 = lContC110 + 1
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "C110")
        Call objEFD.Reg_Campo_add(2, "COD_INF", Format(objDoc.lCodMensagem, "000000"))
        Call objEFD.Reg_Campo_add(3, "TXT_COMPL", "")
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Processa_Registro_C110 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C110:

   Processa_Registro_C110 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209486)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_C111(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC111 As Long) As Long
'REGISTRO C111: PROCESSO REFERENCIADO

On Error GoTo Erro_Processa_Registro_C111

    '******** registro C111 *****************
    objEFD.objSpedFiscal.sTipo = "C111"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_C111 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C111:

   Processa_Registro_C111 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209487)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_C120(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC120 As Long, ByVal lComando As Long) As Long
'REGISTRO C120: COMPLEMENTO DO DOCUMENTO - OPERAÇÕES DE IMPORTAÇÃO (CÓDIGO 01)
'Este registro tem por objetivo informar detalhes das operações de importação, que estejam sendo documentadas pela nota
'fiscal escriturada no registro C100, quando o campo IND_OPER for igual a “0” (zero), indicando operação de entrada e
'que no registro filho C170 conste CST_PIS ou CST_COFINS gerador de crédito (CST 50 a 56), bem como conste no
'registro C170 CFOP próprio de operações de importação (CFOP iniciado em 3).
'Não podem ser informados para um mesmo documento fiscal, dois ou mais registros com o mesmo conteúdo no campo
'NUM_DOC_IMP e NUM_ACDRAW.
'Registro com estrutura, campos e conteúdo definidos e constantes no Leiaute da Escrituração Fiscal Digital – EFD (ICMS
'e IPI), instituído pelo Ato COTEPE/ICMS nº 9, de 12 de abril de 2008, disponível no portal de serviços (SPED) da
'página da Secretaria da Receita Federal do Brasil na Internet, no endereço <www.receita.fazenda.gov.br/sped>.
Dim lErro As Long
Dim tDIInfo As typeDIInfo
Dim sDI As String

On Error GoTo Erro_Processa_Registro_C120

    '******** registro C120 *****************
    objEFD.objSpedFiscal.sTipo = "C120"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    If objDoc.iTipoDoc = SPED_DOC_NF Then
    
        tDIInfo.sNumero = String(STRING_DI_NUMERO, 0)

        'TEM QUE MEXER PORQUE UMA DI PODE ESTAR ASSOCIADA A MAIS DE UMA NF

        'pegar os dados de DIInfo
        lErro = Comando_Executar(lComando, "SELECT DI.Numero, DI.PISValor, DI.COFINSValor FROM ItensNFiscal AS INF, ItemAdicaoDIItemNF AS IAD, ItensADicaoDI AS IDI, AdicaoDI AS ADI, DIINFO AS DI WHERE INF.NumIntNF = ? AND IAD.NumIntItemNF = INF.NumIntDoc AND IDI.NumIntDoc =IAD.NumIntItemAdicaoDI AND ADI.NumIntDoc = IDI.NumIntAdicaoDI AND ADI.NumIntDI = DI.NumIntDoc", _
        tDIInfo.sNumero, tDIInfo.dPISValor, tDIInfo.dCOFINSValor, objDoc.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 209488

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209489

        If lErro = AD_SQL_SUCESSO Then

            lContC120 = lContC120 + 1

            sDI = Replace(Replace(Replace(Replace(Replace(tDIInfo.sNumero, "/", ""), ".", ""), "\", ""), "-", ""), " ", "")

            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "C120")
            Call objEFD.Reg_Campo_add(2, "COD_DOC_IMP", "0")
            Call objEFD.Reg_Campo_add(3, "NUM_DOC_IMP", SFF_Texto(sDI, 10))
            Call objEFD.Reg_Campo_add(4, "VL_PIS_IMP", SFF_Num(tDIInfo.dPISValor))
            Call objEFD.Reg_Campo_add(5, "VL_COFINS_IMP", SFF_Num(tDIInfo.dCOFINSValor))
            Call objEFD.Reg_Campo_add(6, "NUM_ACDRAW", "")
            Call objEFD.Finaliza_Reg

            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If
    
    End If
    
    Processa_Registro_C120 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C120:

   Processa_Registro_C120 = gErr

   Select Case gErr
   
        Case 209488, 209489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209490)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_C170(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long) As Long
'REGISTRO C170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO (CÓDIGOS 01, 1B, 04 e 55)
'Registro obrigatório para discriminar os itens da nota fiscal (mercadorias e/ou serviços constantes em notas conjugadas),
'inclusive em operações de entrada de mercadorias acompanhada de Nota Fiscal Eletrônica (NF-e) de emissão de
'terceiros.
'Não podem ser informados para um mesmo documento fiscal, dois ou mais registros com o mesmo conteúdo no campo
'NUM_ITEM.
'Registro com estrutura, campos e conteúdo definidos e constantes no Leiaute da Escrituração Fiscal Digital – EFD (ICMS
'e IPI), instituído pelo Ato COTEPE/ICMS nº 9, de 12 de abril de 2008, disponível no portal de serviços (SPED) da
'página da Secretaria da Receita Federal do Brasil na Internet, no endereço <www.receita.fazenda.gov.br/sped>.
'IMPORTANTE: para documentos de entrada/aquisição, os campos de valor de imposto/contribuição, base de cálculo e
'alíquota só devem ser informados se o adquirente tiver direito à apropriação do crédito (enfoque do declarante). Não
'precisam ser relacionados documentos fiscais que não dão direito à apuração de créditos de PIS/Pasep e de Cofins. Caso o
'documento fiscal contenha tanto itens sem direito à apropriação de crédito quanto itens com direito, a nota fiscal deverá
'ser informada em sua integralidade.
'Não precisam ser relacionados neste registro os documentos fiscais representativos das operações geradoras de
'contribuição social ou de crédito, abaixo relacionadas, tendo em vista que as mesma são informadas e consideradas em
'registros próprios da EFD-PIS/Cofins:
'- Aquisição de bens a serem incorporados ao ativo imobilizado, cujo crédito for determinado com base no valor
'de aquisição e/ou com base nos encargos mensais de depreciação. O detalhamento do crédito com base nos
'encargos de depreciação deverá ser feito no registro F120. Caso o crédito seja apurado com base no valor de
'aquisição deverá ser informado no registro F130;
'- Fornecimento e/ou Aquisição de Energia Elétrica (documento fiscal código 06). Os documentos fiscais
'relativos à energia elétrica devem ser escriturados nos registros C500 (Aquisição com crédito) e/ou C600
'(Fornecimento de energia);
'- Prestação e/ou Aquisição de serviços de transportes (documentos fiscais códigos 07, 08, 8B, 09, 10, 11, 26, 27
'e 57). Os referidos documentos fiscais relativos a serviços de transportes devem ser escriturados nos registros
'D100 (Aquisição com crédito) e/ou D200 (Prestação de serviço);
'- Prestação e/ou Aquisição de serviços de transporte de passageiros – Bilhetes de Passagem (documentos fiscais
'códigos 2E, 13, 14, 15, 16 e 18). Os referidos documentos fiscais relativos a serviços de transporte de
'passageiros devem ser escriturados nos registros D300 ou D350 (bilhete emitido por ECF);
'- Prestação e/ou Aquisição de serviços de comunicação e telecomunicação (documentos fiscais códigos 21 e
'22). Os referidos documentos fiscais relativos a serviços de comunicação e telecomunicação devem ser
'escriturados nos registros D500 (Aquisição com crédito) e/ou D600 (Prestação de serviço);
'- Fornecimento e/ou Aquisição de água canalizada ou gás (documentos fiscais códigos 28 e 29). Os documentos
'fiscais relativos a água canalizada e gás devem ser escriturados nos registros C500 (Aquisição com crédito) e/ou
'C600 (Fornecimento d´água canalizada e gás);
'- Cupom Fiscal (documentos fiscais códigos 02, 2D e 59). Os documentos fiscais relativos Cupom Fiscal devem
'ser escriturados nos registros C400 (informação por ECF) ou C490 (informação consolidada).
Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim sChave As String

On Error GoTo Erro_Processa_Registro_C170

    '******** registro C170 *****************
    objEFD.objSpedFiscal.sTipo = "C170"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    For Each objItem In objDoc.colItens
    
        sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Item: " & CStr(objItem.iItem)
        Call objEFD.objTela.SPED_Status("Processando Registro C170", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "C170")
        Call objEFD.Reg_Campo_add(2, "NUM_ITEM", Format(objItem.iItem, "000"))
        Call objEFD.Reg_Campo_add(3, "COD_ITEM", SFF_Texto(objItem.sProduto))
        Call objEFD.Reg_Campo_add(4, "DESCR_COMPL", SFF_Texto(objItem.sDescricao))
        Call objEFD.Reg_Campo_add(5, "QTD", SFF_Num(objItem.dQuantidade, 5))
        Call objEFD.Reg_Campo_add(6, "UNID", SFF_Texto(objItem.sUM))
        Call objEFD.Reg_Campo_add(7, "VL_ITEM", SFF_Num(objItem.dPrecoTotal))
        Call objEFD.Reg_Campo_add(8, "VL_DESC", SFF_Num(objItem.dValorDesconto))
        Call objEFD.Reg_Campo_add(9, "IND_MOV", SFF_Texto(objItem.sMovFisica))
        Call objEFD.Reg_Campo_add(10, "CST_ICMS", objItem.sICMSCST)
        Call objEFD.Reg_Campo_add(11, "CFOP", objItem.sCFOP)
        Call objEFD.Reg_Campo_add(12, "COD_NAT", "")
        Call objEFD.Reg_Campo_add(13, "VL_BC_ICMS", SFF_Num(objItem.dValorBaseICMS))
        Call objEFD.Reg_Campo_add(14, "ALIQ_ICMS", SFF_Perc(objItem.dICMSAliquota))
        Call objEFD.Reg_Campo_add(15, "VL_ICMS", SFF_Num(objItem.dValorICMS))
        Call objEFD.Reg_Campo_add(16, "VL_BC_ICMS_ST", SFF_Num(objItem.dValorBaseICMSST))
        Call objEFD.Reg_Campo_add(17, "ALIQ_ST", SFF_Perc(objItem.dICMSSTAliquota))
        Call objEFD.Reg_Campo_add(18, "VL_ICMS_ST", SFF_Num(objItem.dValorICMSST))
        Call objEFD.Reg_Campo_add(19, "IND_APUR", "0")
        Call objEFD.Reg_Campo_add(20, "CST_IPI", objItem.sIPICST)
        Call objEFD.Reg_Campo_add(21, "COD_ENQ", SFF_Texto(objItem.sIPICodEnq))
        Call objEFD.Reg_Campo_add(22, "VL_BC_IPI", SFF_Num(objItem.dValorBaseIPI))
        Call objEFD.Reg_Campo_add(23, "ALIQ_IPI", SFF_Perc(objItem.dIPIAliquota))
        Call objEFD.Reg_Campo_add(24, "VL_IPI", SFF_Num(objItem.dValorIPI))
        Call objEFD.Reg_Campo_add(25, "CST_PIS", objItem.sPISCST)
        Call objEFD.Reg_Campo_add(26, "VL_BC_PIS", IIf(Abs(objItem.dPISBCQtd) > QTDE_ESTOQUE_DELTA, "", SFF_Num(objItem.dValorBasePIS)))
        Call objEFD.Reg_Campo_add(27, "ALIQ_PIS", IIf(Abs(objItem.dPISBCQtd) > QTDE_ESTOQUE_DELTA, "", SFF_Perc(objItem.dPISAliquota, 4)))
        Call objEFD.Reg_Campo_add(28, "QUANT_BC_PIS", IIf(Abs(objItem.dPISBCQtd) > QTDE_ESTOQUE_DELTA, SFF_Num(objItem.dPISBCQtd, 3), ""))
        Call objEFD.Reg_Campo_add(29, "ALIQ_PIS_QUANT", IIf(Abs(objItem.dPISBCQtd) > QTDE_ESTOQUE_DELTA, SFF_Num(objItem.dPISAliquotaQtd, 4), ""))
        Call objEFD.Reg_Campo_add(30, "VL_PIS", SFF_Num(objItem.dValorPIS))
        Call objEFD.Reg_Campo_add(31, "CST_COFINS", objItem.sCofinsCST)
        Call objEFD.Reg_Campo_add(32, "VL_BC_COFINS", IIf(Abs(objItem.dCofinsBCQtd) > QTDE_ESTOQUE_DELTA, "", SFF_Num(objItem.dValorBaseCofins)))
        Call objEFD.Reg_Campo_add(33, "ALIQ_COFINS", IIf(Abs(objItem.dCofinsBCQtd) > QTDE_ESTOQUE_DELTA, "", SFF_Perc(objItem.dCOFINSAliquota, 4)))
        Call objEFD.Reg_Campo_add(34, "QUANT_BC_COFINS", IIf(Abs(objItem.dCofinsBCQtd) > QTDE_ESTOQUE_DELTA, SFF_Num(objItem.dCofinsBCQtd, 3), ""))
        Call objEFD.Reg_Campo_add(35, "ALIQ_COFINS_QUANT", IIf(Abs(objItem.dCofinsBCQtd) > QTDE_ESTOQUE_DELTA, SFF_Num(objItem.dCofinsAliquotaQtd, 4), ""))
        Call objEFD.Reg_Campo_add(36, "VL_COFINS", SFF_Num(objItem.dValorCofins))
        Call objEFD.Reg_Campo_add(37, "COD_CTA", "")
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
    
    Processa_Registro_C170 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C170:

   Processa_Registro_C170 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209491)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C170", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C175(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long) As Long
'REGISTRO C175: REGISTRO ANALÍTICO DO DOCUMENTO (CÓDIGO 65)
'Este registro tem por objetivo representar a escrituração da NFC-e, código 65, os documentos fiscais
'totalizados por CST PIS,CST Cofins, CFOP, alíquota de PIS e alíquota da Cofins. Trata-se de registro
'com procedimento de escrituração similar aoadotado para o registro C190 da EFD-ICMS/IPI.
'Atenção:
'1. Registro analítico das receitas decorrentes de emissão de NFC-e, disponível na versão 2.09 do PVA da
'EFD-Contribuições,para utilização na escrituração dos fatos geradores a partir de setembro de 2014.
'2. As receitas eventualmente auferidas com a emissão de NFC-e, referentes aos períodos anteriores ao da
'disponibilidade da versão 2.09, devem ser escrituradas de forma consolidada no Registro C180, devendo
'observar as instruções de preenchimento do referido registro.
'3. Apesar da escrituração do documento ser de forma individualizada por meio do registro C100, a
'escrituração da NFC-e referente ao seu conteúdo, seu itens, não é feita de forma individualizada
'item a item (no registro C170, como é o procedimento para a NF-e, código 55), mas sim, de forma
'analítica, neste registro C175.
'4. Para tanto, será gerado um registro C175, para consolidar todos os itens do documento que tenham o
'mesmo CST e Alíquota de PIS/Pasep e Cofins.
'Exemplo:
'Considerando que determinada NFC-e emitida, contendo dezenas de itens/produtos diferentes em seu conteúdo
'apresente, por exemplo, 18 itens sujeitos à tributação às alíquotas básicas (CST 01) + 30 itens referentes
'a revenda de produtos sujeitos à alíquota zero (CST 06) + 05 cinco itens referentes a revenda de produtos
'sujeitos à substituição tributária de PIS/Pasep e Cofins (CST 05) + 30 itens referentes a revenda de produtos
'sujeitos à tributação monofásica (CST 04), deve então a pessoa jurídica escriturar este documento em
'01 (um) registro C100 e detalhar os valores dos itens contidos na NFC-e em 04 (quatro) registros C175
', consolidando em cada registro C175 as receitas referentes ao cada CST.
'Validação do Registro:
'não podem ser informados dois ou mais registros com a mesma combinação de valores dos campos:
'CFOP, CST (PIS/Pasep e Cofins) e alíquotas (PIS/Pasep e Cofins)
Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim objItemAux As ClassSpedDocFiscaisItens
Dim sChave As String, colItemAgrupado As New Collection
Dim bAchou As Boolean

On Error GoTo Erro_Processa_Registro_C175

    '******** registro C175 *****************
    objEFD.objSpedFiscal.sTipo = "C175"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    For Each objItem In objDoc.colItens
        bAchou = False
        For Each objItemAux In colItemAgrupado
            If objItemAux.sCFOP = objItem.sCFOP And _
            objItemAux.sPISCST = objItem.sPISCST And _
            objItemAux.sCofinsCST = objItem.sCofinsCST And _
            Abs(objItemAux.dPISAliquota - objItem.dPISAliquota) < DELTA_VALORMONETARIO And _
            Abs(objItemAux.dPISAliquotaQtd - objItem.dPISAliquotaQtd) < DELTA_VALORMONETARIO And _
            Abs(objItemAux.dCOFINSAliquota - objItem.dCOFINSAliquota) < DELTA_VALORMONETARIO And _
            Abs(objItemAux.dCofinsAliquotaQtd - objItem.dCofinsAliquotaQtd) < DELTA_VALORMONETARIO Then
                
                'Totaliza os valores por CFOP, CST (PIS/Pasep e Cofins) e alíquotas (PIS/Pasep e Cofins)
                objItemAux.dPrecoTotal = objItemAux.dPrecoTotal + objItem.dPrecoTotal
                objItemAux.dValorDesconto = objItemAux.dValorDesconto + objItem.dValorDesconto
                
                objItemAux.dValorBasePIS = objItemAux.dValorBasePIS + objItem.dValorBasePIS
                objItemAux.dPISBCQtd = objItemAux.dPISBCQtd + objItem.dPISBCQtd
                objItemAux.dValorPIS = objItemAux.dValorPIS + objItem.dValorPIS
                
                objItemAux.dValorBaseCofins = objItemAux.dValorBaseCofins + objItem.dValorBaseCofins
                objItemAux.dCofinsBCQtd = objItemAux.dCofinsBCQtd + objItem.dCofinsBCQtd
                objItemAux.dValorCofins = objItemAux.dValorCofins + objItem.dValorCofins
                
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemAux = New ClassSpedDocFiscaisItens
            
            objItemAux.sCFOP = objItem.sCFOP
            objItemAux.dPrecoTotal = objItem.dPrecoTotal
            objItemAux.dValorDesconto = objItem.dValorDesconto
            objItemAux.sPISCST = objItem.sPISCST
            objItemAux.dValorBasePIS = objItem.dValorBasePIS
            objItemAux.dPISAliquota = objItem.dPISAliquota
            objItemAux.dPISBCQtd = objItem.dPISBCQtd
            objItemAux.dPISAliquotaQtd = objItem.dPISAliquotaQtd
            objItemAux.dValorPIS = objItem.dValorPIS
            objItemAux.sCofinsCST = objItem.sCofinsCST
            objItemAux.dValorBaseCofins = objItem.dValorBaseCofins
            objItemAux.dCOFINSAliquota = objItem.dCOFINSAliquota
            objItemAux.dCofinsBCQtd = objItem.dCofinsBCQtd
            objItemAux.dCofinsAliquotaQtd = objItem.dCofinsAliquotaQtd
            objItemAux.dValorCofins = objItem.dValorCofins
            
            colItemAgrupado.Add objItemAux
        End If
    Next
    
    For Each objItem In colItemAgrupado
    
        sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Item: " & CStr(objItem.iItem)
        Call objEFD.objTela.SPED_Status("Processando Registro C175", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "C175")
        Call objEFD.Reg_Campo_add(2, "CFOP", objItem.sCFOP)
        Call objEFD.Reg_Campo_add(3, "VL_OPR", SFF_Num(objItem.dPrecoTotal))
        Call objEFD.Reg_Campo_add(4, "VL_DESC", SFF_Num(objItem.dValorDesconto))
        Call objEFD.Reg_Campo_add(5, "CST_PIS", objItem.sPISCST)
        Call objEFD.Reg_Campo_add(6, "VL_BC_PIS", SFF_Num(objItem.dValorBasePIS))
        Call objEFD.Reg_Campo_add(7, "ALIQ_PIS", SFF_Perc(objItem.dPISAliquota, 4))
        Call objEFD.Reg_Campo_add(8, "QUANT_BC_PIS", IIf(objItem.dPISBCQtd <> 0, SFF_Num(objItem.dPISBCQtd, 3), ""))
        Call objEFD.Reg_Campo_add(9, "ALIQ_PIS_QUANT", IIf(objItem.dPISAliquotaQtd <> 0, SFF_Num(objItem.dPISAliquotaQtd, 4), ""))
        Call objEFD.Reg_Campo_add(10, "VL_PIS", SFF_Num(objItem.dValorPIS))
        Call objEFD.Reg_Campo_add(11, "CST_COFINS", objItem.sCofinsCST)
        Call objEFD.Reg_Campo_add(12, "VL_BC_COFINS", SFF_Num(objItem.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(13, "ALIQ_COFINS", SFF_Perc(objItem.dCOFINSAliquota, 4))
        Call objEFD.Reg_Campo_add(14, "QUANT_BC_COFINS", IIf(objItem.dCofinsBCQtd <> 0, SFF_Num(objItem.dCofinsBCQtd, 3), ""))
        Call objEFD.Reg_Campo_add(15, "ALIQ_COFINS_QUANT", IIf(objItem.dCofinsAliquotaQtd <> 0, SFF_Num(objItem.dCofinsAliquotaQtd, 4), ""))
        Call objEFD.Reg_Campo_add(16, "VL_COFINS", SFF_Num(objItem.dValorCofins))
        Call objEFD.Reg_Campo_add(17, "COD_CTA", "")
        Call objEFD.Reg_Campo_add(18, "INFO_COMPL", "")
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
    
    Processa_Registro_C175 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C175:

   Processa_Registro_C175 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209491)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C170", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C500(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC501 As Long, lContC505 As Long, lContC509 As Long) As Long
'REGISTRO C500: NOTA FISCAL/CONTA DE ENERGIA ELÉTRICA (CÓDIGO 06), NOTA FISCAL/CONTA DE FORNECIMENTO D'ÁGUA CANALIZADA
'(CÓDIGO 29) E NOTA FISCAL CONSUMO FORNECIMENTO DE GÁS (CÓDIGO 28) - DOCUMENTOS DE ENTRADA/AQUISIÇÃO COM CRÉDITO
'Neste registro serão informadas pela pessoa jurídica as operações sujeitas à apuração de créditos de PIS/Pasep e de Cofins,
'na forma da legislação tributária, referentes a:
'- energia elétrica, consumida nos estabelecimentos da pessoa jurídica (art. 3º, III, das Leis nº 10.637/02 e nº 10.833/03);
'- água canalizada ou gás,  utilizados como insumo na fabricação de produtos destinados à venda ou na prestação de serviços
'(art. 3º, II, das Leis nº 10.637/02 e nº 10.833/03).
'Os documentos fiscais escriturados nestes itens não devem ser relacionados ou escriturados nos Registros C100 ou C190.

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String

On Error GoTo Erro_Processa_Registro_C500

    '******** registro C500 *****************
    objEFD.objSpedFiscal.sTipo = "C500"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro C500", sChave)
    
    If objDoc.lCliente > 0 Then
        sCodParticipante = "C" & Format(objDoc.lCliente, "0000000000") & Format(objDoc.iFilialCli, "00000")
    Else
        sCodParticipante = "F" & Format(objDoc.lFornecedor, "0000000000") & Format(objDoc.iFilialForn, "00000")
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "C500")
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
            Call objEFD.Reg_Campo_add(2, "COD_PART", "")
        
        Case Else
            Call objEFD.Reg_Campo_add(2, "COD_PART", sCodParticipante)
            
    End Select
    
    Call objEFD.Reg_Campo_add(3, "COD_MOD", objDoc.sModelo)
    Call objEFD.Reg_Campo_add(4, "COD_SIT", Format(objDoc.iSituacao, "00"))
    Call objEFD.Reg_Campo_add(5, "SER", SFF_Texto(objDoc.sSerie))
    Call objEFD.Reg_Campo_add(6, "SUB", "")
    Call objEFD.Reg_Campo_add(7, "NUM_DOC", CStr(objDoc.lNumDocumento))
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
            Call objEFD.Reg_Campo_add(8, "DT_DOC", "")
            Call objEFD.Reg_Campo_add(9, "DT_ENT", "")
            Call objEFD.Reg_Campo_add(10, "VL_DOC", "")
            Call objEFD.Reg_Campo_add(11, "VL_ICMS", "")
            Call objEFD.Reg_Campo_add(12, "COD_INF", "")
            Call objEFD.Reg_Campo_add(13, "VL_PIS", "")
            Call objEFD.Reg_Campo_add(14, "VL_COFINS", "")
        
        Case Else
        
            Call objEFD.Reg_Campo_add(8, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
            Call objEFD.Reg_Campo_add(9, "DT_ENT", SFF_Data(objDoc.dtDataES))
            Call objEFD.Reg_Campo_add(10, "VL_DOC", SFF_Num(objDoc.dValorTotal))
            Call objEFD.Reg_Campo_add(11, "VL_ICMS", SFF_Num(objDoc.dValorICMS))
            If objDoc.lCodMensagem > 0 Then
                Call objEFD.Reg_Campo_add(12, "COD_INF", Format(objDoc.lCodMensagem, "000000"))
            Else
                Call objEFD.Reg_Campo_add(12, "COD_INF", "")
            End If
            Call objEFD.Reg_Campo_add(13, "VL_PIS", SFF_Num(objDoc.dValorPIS))
            Call objEFD.Reg_Campo_add(14, "VL_COFINS", SFF_Num(objDoc.dValorCofins))
            
    End Select
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
        Case Else
            
            lErro = Processa_Registro_C501(objEFD, objDoc, lSeq, lContC501)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            lErro = Processa_Registro_C505(objEFD, objDoc, lSeq, lContC505)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Processa_Registro_C509(objEFD, objDoc, lSeq, lContC509)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End Select
    
    Processa_Registro_C500 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C500:

   Processa_Registro_C500 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211066)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C500", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C501(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC501 As Long) As Long
'REGISTRO C501: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 06, 28 e 29) - PIS/PASEP
'Neste registro devem ser detalhadas as informações relativas à apuração do crédito de PIS/Pasep, referentes ao
'documento fiscal escriturado no Registro Pai C500. Deve ser escriturado um registro C501 para cada item
'(fornecimento dágua canalizada, de energia elétrica ou de gás) cuja operação dê direito a crédito, pelo seu valor
'total ou parcial.
'Caso em relação a um mesmo item venha a ocorrer tratamentos tributários diversos (mais de um CST), deve a pessoa jurídica
'informar um registro C501 para cada CST.

Dim lErro As Long
Dim sCodParticipante As String
Dim sChave As String
Dim objItem As ClassSpedDocFiscaisItens
Dim colAux As New Collection, bAchou As Boolean
Dim objItemAux As ClassSpedDocFiscaisItens

On Error GoTo Erro_Processa_Registro_C501

    '******** registro C501 *****************
    objEFD.objSpedFiscal.sTipo = "C501"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D501", sChave)
    
    For Each objItem In objDoc.colItens
        bAchou = False
        For Each objItemAux In colAux
            If objItemAux.sPISCST = objItem.sPISCST Then
                objItemAux.dValorBasePIS = objItemAux.dValorBasePIS + objItem.dValorBasePIS
                objItemAux.dValorPIS = objItemAux.dValorPIS + objItem.dValorPIS
                objItemAux.dPrecoTotal = objItemAux.dPrecoTotal + objItem.dPrecoTotal
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemAux = New ClassSpedDocFiscaisItens
            objItemAux.sPISCST = objItem.sPISCST
            objItemAux.dValorBasePIS = objItem.dValorBasePIS
            objItemAux.dValorPIS = objItem.dValorPIS
            objItemAux.dPrecoTotal = objItem.dPrecoTotal
            objItemAux.dPISAliquota = objItem.dPISAliquota
            objItemAux.sNatBCCred = objItem.sNatBCCred
            colAux.Add objItemAux
        End If
    Next
    
    For Each objItemAux In colAux
        
        lContC501 = lContC501 + 1
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "C501")
        Call objEFD.Reg_Campo_add(2, "CST_PIS", objItemAux.sPISCST)
        Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objItemAux.dPrecoTotal))
        Call objEFD.Reg_Campo_add(4, "NAT_BC_CRED", SFF_Texto(objItemAux.sNatBCCred))
        Call objEFD.Reg_Campo_add(5, "VL_BC_PIS", SFF_Num(objItemAux.dValorBasePIS))
        Call objEFD.Reg_Campo_add(6, "ALIQ_PIS", SFF_Perc(objItemAux.dPISAliquota))
        Call objEFD.Reg_Campo_add(7, "VL_PIS", SFF_Num(objItemAux.dValorPIS))
        Call objEFD.Reg_Campo_add(8, "COD_CTA", "")
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
       
    Processa_Registro_C501 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C501:

   Processa_Registro_C501 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211067)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C501", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C505(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC505 As Long) As Long
'REGISTRO C505: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 06, 28 e 29) - COFINS
'Neste registro devem ser detalhadas as informações relativas à apuração do crédito de COFINS,
'referentes ao documento fiscal escriturado no Registro Pai C500. Deve ser escriturado um registro C505 para cada
'item (fornecimento dágua canalizada, de energia elétrica ou de gás) cuja operação dê direito a crédito, pelo seu valor
'total ou parcial.
'Caso em relação a um mesmo item venha a ocorrer tratamentos tributários diversos (mais de um CST), deve a pessoa jurídica
'informar um registro C505 para cada CST.

Dim lErro As Long
Dim sCodParticipante As String
Dim sChave As String
Dim objItem As ClassSpedDocFiscaisItens
Dim colAux As New Collection, bAchou As Boolean
Dim objItemAux As ClassSpedDocFiscaisItens

On Error GoTo Erro_Processa_Registro_C505

    '******** registro C505 *****************
    objEFD.objSpedFiscal.sTipo = "C505"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro C505", sChave)
    
    For Each objItem In objDoc.colItens
        bAchou = False
        For Each objItemAux In colAux
            If objItemAux.sCofinsCST = objItem.sCofinsCST Then
                objItemAux.dValorBaseCofins = objItemAux.dValorBaseCofins + objItem.dValorBaseCofins
                objItemAux.dValorCofins = objItemAux.dValorCofins + objItem.dValorCofins
                objItemAux.dPrecoTotal = objItemAux.dPrecoTotal + objItem.dPrecoTotal
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemAux = New ClassSpedDocFiscaisItens
            objItemAux.sCofinsCST = objItem.sCofinsCST
            objItemAux.dValorBaseCofins = objItem.dValorBaseCofins
            objItemAux.dValorCofins = objItem.dValorCofins
            objItemAux.dPrecoTotal = objItem.dPrecoTotal
            objItemAux.dCOFINSAliquota = objItem.dCOFINSAliquota
            objItemAux.sNatBCCred = objItem.sNatBCCred
            colAux.Add objItemAux
        End If
    Next
    
    For Each objItemAux In colAux
        
        lContC505 = lContC505 + 1
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "C505")
        Call objEFD.Reg_Campo_add(2, "CST_COFINS", objItemAux.sCofinsCST)
        Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objItemAux.dPrecoTotal))
        Call objEFD.Reg_Campo_add(4, "NAT_BC_CRED", SFF_Texto(objItemAux.sNatBCCred))
        Call objEFD.Reg_Campo_add(5, "VL_BC_COFINS", SFF_Num(objItemAux.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(6, "ALIQ_COFINS", SFF_Perc(objItemAux.dCOFINSAliquota))
        Call objEFD.Reg_Campo_add(7, "VL_COFINS", SFF_Num(objItemAux.dValorCofins))
        Call objEFD.Reg_Campo_add(8, "COD_CTA", "")
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
       
    Processa_Registro_C505 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C505:

   Processa_Registro_C505 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211068)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_C505", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_C509(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContC509 As Long) As Long
'REGISTRO C509: PROCESSO REFERENCIADO

On Error GoTo Erro_Processa_Registro_C509

    '******** registro C509 *****************
    objEFD.objSpedFiscal.sTipo = "C509"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_C509 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C509:

   Processa_Registro_C509 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211069)
    
    End Select
        
    Exit Function
    
End Function

Private Function Processa_Registro_C990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO C990: ENCERRAMENTO DO BLOCO C
'Este registro destina-se a identificar o encerramento do bloco C e informar a quantidade de linhas (registros) existentes no
'bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_C990

    '******************* registro C990 **********************************
    objEFD.objSpedFiscal.sTipo = "C990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1 ' ele mesmo
    
    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "C" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "C990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_C", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_C990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_C990:

   Processa_Registro_C990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209492)
    
    End Select
    
    Exit Function
    
End Function












Private Function Processa_Registro_9001(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 9001: ABERTURA DO BLOCO 9
'Este registro deve sempre ser gerado e representa a abertura do Bloco 9.
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_9001

    '******************* registro 9001 **********************************
    objEFD.objSpedFiscal.sTipo = "9001"
    objEFD.objSpedFiscal.lSeqPai = 0

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "9001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", "0")
    Call objEFD.Finaliza_Reg
        
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9001:

   Processa_Registro_9001 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209493)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9900(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 9900: REGISTROS DO ARQUIVO
'Todos os registros referenciados neste arquivo, inclusive os posteriores a este registro, devem ter uma linha totalizadora
'do seu número de ocorrências.
Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9900

    '******************* registro 9900 **********************************
    objEFD.objSpedFiscal.sTipo = "9900"
    objEFD.objSpedFiscal.lSeqPai = 0

    Call Registro_Adiciona_Count("9990", objEFD.colSpedRegistro, 1)
    Call Registro_Adiciona_Count("9999", objEFD.colSpedRegistro, 1)
    Call Registro_Adiciona_Count(objEFD.objSpedFiscal.sTipo, objEFD.colSpedRegistro, objEFD.colSpedRegistro.Count + 1)

    For Each objSpedRegistro In objEFD.colSpedRegistro

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "9900")
        Call objEFD.Reg_Campo_add(2, "REG_BLC", objSpedRegistro.sTipo)
        Call objEFD.Reg_Campo_add(3, "QTD_REG_BLC", CStr(objSpedRegistro.lCount))
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next

    Processa_Registro_9900 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9900:

   Processa_Registro_9900 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209494)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
'Este registro destina-se a identificar o encerramento do Bloco 9 e a informar a quantidade de linhas (registros) existentes
'no bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9990

    '******************* registro 9990 **********************************
    objEFD.objSpedFiscal.sTipo = "9990"
    objEFD.objSpedFiscal.lSeqPai = 0

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "9" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "9990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_9", CStr(lContador))
    Call objEFD.Finaliza_Reg
            
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9990:

   Processa_Registro_9990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209495)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_9999(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
'Este registro destina-se a identificar o encerramento do arquivo digital da escrituração do PIS/Pasep e da Cofins, bem
'como a informar a quantidade de linhas (registros) existentes no arquivo.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_9999

    '******************* registro 9999 **********************************
    objEFD.objSpedFiscal.sTipo = "9999"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    For Each objSpedRegistro In objEFD.colSpedRegistro
        lContador = lContador + objSpedRegistro.lCount
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "9999")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN", CStr(lContador))
    Call objEFD.Finaliza_Reg

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_9999 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_9999:

   Processa_Registro_9999 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209496)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_D001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBloco As Boolean) As Long
'REGISTRO D001: ABERTURA DO BLOCO D
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_D001

    '******** registro D001 *****************

    'Se  for a versão inferior a 3 ou a filial não seja lucro presumido ou não cumulativo ou sem regime de caixa, lê doc a doc
    If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA Then
        
        lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE  NumIntSped = ? AND Bloco = ?", lContador, objEFD.lNumIntSped, "D")
        If lErro <> AD_SQL_SUCESSO Then gError 201273
    
        lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201274
        
    End If
        
    objEFD.objSpedFiscal.sTipo = "D001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBloco = True
    Else
        iTemReg = 1
        bAbreBloco = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_D001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D001:

   Processa_Registro_D001 = gErr

   Select Case gErr
   
        Case 201273, 201274
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201275)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_D010(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO D010: IDENTIFICAÇÃO DO ESTABELECIMENTO
'Este registro tem o objetivo de identificar o estabelecimento da pessoa jurídica a que se referem as operações e
'documentos fiscais informados neste bloco. Só devem ser escriturados no Registro D010 os estabelecimentos que
'efetivamente tenham realizado as operações especificadas no Bloco D (prestação ou contratação), relativas a serviços de
'transporte de cargas e/ou de passageiros, serviços de comunicação e de telecomunicação, mediante emissão de documento
'fiscal definido pela legislação do ICMS e do IPI, que devam ser escrituradas no Bloco D.
'O estabelecimento que não realizou operações passíveis de registro nesse bloco, no período da escrituração, não deve ser
'identificado no Registro D010.
'Para cada estabelecimento cadastrado em “D010”, deve ser informado nos registros de nível inferior (Registros Filho) as
'operações próprias de prestação ou de contratação, mediante emissão de documento fiscal, no mercado interno ou externo

Dim lErro As Long, iIndice As Integer
Dim lContD010 As Long, lContD100 As Long, lContD101 As Long, lContD105 As Long, lContD111 As Long
Dim lContD200 As Long, lContD205 As Long, lContD209 As Long
Dim lContD500 As Long, lContD501 As Long, lContD505 As Long, lContD509 As Long

Dim objDoc As ClassSpedDocFiscais
Dim objDocItens As ClassSpedDocFiscaisItens
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 3) As Long
Dim sCGCAnt As String, sChave As String

On Error GoTo Erro_Processa_Registro_D010

    '******************* registro D010 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    sCGCAnt = "X"  'Força ser diferente
    
    lErro = SpedFiscal_Docs_Prepara(alComando(0), objEFD.lNumIntSped, "D", tDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201276
   
    Do While lErro = AD_SQL_SUCESSO
        
        Set objDoc = New ClassSpedDocFiscais
        
        lErro = SpedFiscal_Doc_Copia(objDoc, tDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = SpedFiscal_DocItens_Le(alComando(1), objDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Mudou o CPF\CNPJ grava o D010
        If sCGCAnt <> objDoc.sCgc Then
            
            sChave = "CNPJ\CPF: " & objDoc.sCgc
            Call objEFD.objTela.SPED_Status("Processando Registro D010", sChave)
            
            objEFD.objSpedFiscal.sTipo = "D010"
            objEFD.objSpedFiscal.lSeqPai = 0
            
            lContD010 = lContD010 + 1
            
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "D010")
            Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(objDoc.sCgc, 14))
            Call objEFD.Finaliza_Reg
            
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sCGCAnt = objDoc.sCgc
        End If
                
        Select Case objDoc.sRegistro
        
            Case "D100" 'visão documental

                lContD100 = lContD100 + 1

                lErro = Processa_Registro_D100(objEFD, objDoc, objEFD.objSpedFiscal.lSeq, lContD101, lContD105, lContD111)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case "D200" 'Prestação de Serviços de Transportes ... (não temos clientes nesse perfil, não vai ser feito por enquanto)

            Case "D300" 'BILHETES CONSOLIDADOS DE PASSAGEM RODOVIÁRIO (CÓDIGO 13), DE PASSAGEM AQUAVIÁRIO (CÓDIGO 14), DE PASSAGEM E NOTA DE BAGAGEM (CÓDIGO 15), DE PASSAGEM FERROVIÁRIO (CÓDIGO 16) E RESUMO DE MOVIMENTO DIÁRIO (CÓDIGO 18) ... (não temos clientes nesse perfil, não vai ser feito por enquanto)

            Case "D350" 'CUPOM FISCAL EMITIDO POR ECF - (CÓDIGOS 2E, 13, 14, 15 e 16) ... (não temos clientes nesse perfil, não vai ser feito por enquanto)

            Case "D500" 'Aquisição de de Serviços - Notas de Serviço de Comunicação (Código 21) e de Serviço de Telecomunicação (Código 22)

                lContD500 = lContD500 + 1

                lErro = Processa_Registro_D500(objEFD, objDoc, objEFD.objSpedFiscal.lSeq, lContD501, lContD505, lContD509)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case "D600" 'Prestação de Serviços - Notas de Serviço de Comunicação (Código 21) e de Serviço de Telecomunicação (Código 22) ... (não temos clientes nesse perfil, não vai ser feito por enquanto)
            
        End Select
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201277
    
    Loop

    Processa_Registro_D010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D010:

   Processa_Registro_D010 = gErr

   Select Case gErr
   
        Case 201276, 201277
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209484)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D010", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D100(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD101 As Long, lContD105 As Long, lContD111 As Long) As Long
'REGISTRO D100: AQUISIÇÃO DE SERVIÇOS DE TRANSPORTE - NOTA FISCAL DE SERVIÇO DE TRANSPORTE (CÓDIGO 07) E CONHECIMENTOS
'DE TRANSPORTE RODOVIÁRIO DE CARGAS (CÓDIGO 08), CONHECIMENTO DE TRANSPORTE DE CARGAS AVULSO (CÓDIGO 8B), AQUAVIÁRIO DE
'CARGAS (CÓDIGO 09), AÉREO (CÓDIGO 10), FERROVIÁRIO DE CARGAS (CÓDIGO 11), MULTIMODAL DE CARGAS (CÓDIGO 26), NOTA FISCAL
'DE TRANSPORTE FERROVIÁRIO DE CARGA (CÓDIGO 27) E CONHECIMENTO DE TRANSPORTE ELETRÔNICO - CT-e (CÓDIGO 57)
'Este registro deve ser apresentado por todos os contribuintes adquirentes dos serviços relacionados, que utilizem os
'documentos previstos para este registro, cuja operação dê direito à apuração de crédito à pessoa jurídica contratante,
'na forma da legislação tributária.
'1. As seguintes operações de transportes dão direito a crédito, de acordo com a legislação e atos normativos aplicáveis
'ao PIS/Pasep e à Cofins:
'- Fretes incorridos nas operações de revenda de mercadorias e produtos, quando o ônus for suportado pela pessoa
'jurídica comercial titular da escrituração (contratação de frete para a entrega da mercadoria revendida ao adquirente);
'- Fretes incorridos nas operações de venda de bens e produtos fabricados a pessoa jurídica titular da escrituração,
'quando o ônus for suportado pela pessoa jurídica titular da escrituração (contratação de frete para a entrega de bens e
'produtos vendidos ao adquirente).
'OBS: Os gastos com transporte na aquisição das mercadorias podem compor a base de cálculo dos créditos não cumulativos,
'uma vez que consoante a boa técnica contábil e a legislação fiscal (art. 289, § 1º, do RIR/1999) integra o custo de
'aquisição das mercadorias adquiridas, o frete, quando pago pela pessoa jurídica adquirente. O valor do frete pago
'pela pessoa jurídica na aquisição de mercadorias pode, assim, compor a base de cálculo do crédito referente às
'aquisições dos bens objeto de informação em C100 (escrituração por documento fiscal) ou em C190 (escrituração consolidada).
'2. As seguintes operações de transportes não estão relacionadas na legislação e atos normativos aplicáveis ao PIS/Pasep
'e à Cofins, como operações com direito à apuração de crédito:
'- Os gastos com transporte do produto, acabado ou em elaboração, entre estabelecimentos industriais ou distribuidores
'da mesma pessoa jurídica (transferências de mercadorias e produtos);
'- O transporte de bens recebidos em devolução, realizado do estabelecimento do comprador para o do vendedor.
'
'Validação do Registro: não podem ser informados dois ou mais registros com a combinação de mesmos valores dos campos :
'1. emissão de terceiros : IND_EMIT+NUM_DOC+COD_MOD+SER+SUB+COD_PART;
'2. emissão própria: IND_EMIT+NUM_DOC+COD_MOD+SER+SUB.
'Para cada documento informado e relacionado em cada registro D100, obrigatoriamente deve ser apresentado o detalhamento
'das informações, por item do documento, referentes ao PIS/Pasep (D101) e à Cofins (D105).

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String

On Error GoTo Erro_Processa_Registro_D100

    '******** registro D100 *****************
    objEFD.objSpedFiscal.sTipo = "D100"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D100", sChave)
    
    If objDoc.lCliente > 0 Then
        sCodParticipante = "C" & Format(objDoc.lCliente, "0000000000") & Format(objDoc.iFilialCli, "00000")
    Else
        sCodParticipante = "F" & Format(objDoc.lFornecedor, "0000000000") & Format(objDoc.iFilialForn, "00000")
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D100")
    Call objEFD.Reg_Campo_add(2, "IND_OPER", objDoc.sOperacao)
    Call objEFD.Reg_Campo_add(3, "IND_EMIT", objDoc.sEmitente)
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
            Call objEFD.Reg_Campo_add(4, "COD_PART", "")
        
        Case Else
            Call objEFD.Reg_Campo_add(4, "COD_PART", sCodParticipante)
            
    End Select
    
    Call objEFD.Reg_Campo_add(5, "COD_MOD", objDoc.sModelo)
    Call objEFD.Reg_Campo_add(6, "COD_SIT", Format(objDoc.iSituacao, "00"))
    Call objEFD.Reg_Campo_add(7, "SER", SFF_Texto(objDoc.sSerie))
    Call objEFD.Reg_Campo_add(8, "SUB", "")
    Call objEFD.Reg_Campo_add(9, "NUM_DOC", CStr(objDoc.lNumDocumento))
    Call objEFD.Reg_Campo_add(10, "CHV_CTE", objDoc.sChaveEletronica)
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
            Call objEFD.Reg_Campo_add(11, "DT_DOC", "")
            Call objEFD.Reg_Campo_add(12, "DT_A_P", "")
            Call objEFD.Reg_Campo_add(13, "TP_CT-e", "")
            Call objEFD.Reg_Campo_add(14, "CHV_CTE_REF", "")
            Call objEFD.Reg_Campo_add(15, "VL_DOC", "")
            Call objEFD.Reg_Campo_add(16, "VL_DESC", "")
            Call objEFD.Reg_Campo_add(17, "IND_FRT", "")
            Call objEFD.Reg_Campo_add(18, "VL_SERV", "")
            Call objEFD.Reg_Campo_add(19, "VL_BC_ICMS", "")
            Call objEFD.Reg_Campo_add(20, "VL_ICMS", "")
            Call objEFD.Reg_Campo_add(21, "VL_NT", "")
            Call objEFD.Reg_Campo_add(22, "COD_INF", "")
            Call objEFD.Reg_Campo_add(23, "COD_CTA", "")
        
        Case Else
        
            Call objEFD.Reg_Campo_add(11, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
            Call objEFD.Reg_Campo_add(12, "DT_A_P", SFF_Data(objDoc.dtDataES))
            Call objEFD.Reg_Campo_add(13, "TP_CT-e", "")
            Call objEFD.Reg_Campo_add(14, "CHV_CTE_REF", "")
            Call objEFD.Reg_Campo_add(15, "VL_DOC", SFF_Num(objDoc.dValorTotal))
            Call objEFD.Reg_Campo_add(16, "VL_DESC", SFF_Num(objDoc.dValorDesconto))
            Call objEFD.Reg_Campo_add(17, "IND_FRT", objDoc.sFrete)
            Call objEFD.Reg_Campo_add(18, "VL_SERV", SFF_Num(objDoc.dValorServico))
            Call objEFD.Reg_Campo_add(19, "VL_BC_ICMS", SFF_Num(objDoc.dValorBaseICMS))
            Call objEFD.Reg_Campo_add(20, "VL_ICMS", SFF_Num(objDoc.dValorICMS))
            Call objEFD.Reg_Campo_add(21, "VL_NT", SFF_Num(objDoc.dValorNaoTributado))
            If objDoc.lCodMensagem > 0 Then
                Call objEFD.Reg_Campo_add(22, "COD_INF", Format(objDoc.lCodMensagem, "000000"))
            Else
                Call objEFD.Reg_Campo_add(22, "COD_INF", "")
            End If
            Call objEFD.Reg_Campo_add(23, "COD_CTA", "")
    
    End Select
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
        Case Else
            
            lErro = Processa_Registro_D101(objEFD, objDoc, lSeq, lContD101)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            lErro = Processa_Registro_D105(objEFD, objDoc, lSeq, lContD105)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Processa_Registro_D111(objEFD, objDoc, lSeq, lContD111)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End Select
    
    Processa_Registro_D100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D100:

   Processa_Registro_D100 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211071)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D100", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D101(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD101 As Long) As Long
'REGISTRO D101: COMPLEMENTO DO DOCUMENTO DE TRANSPORTE (Códigos 07, 08, 8B, 09, 10, 11, 26, 27 e 57) - PIS/PASEP
'Serão escrituradas neste registro as informações referentes à incidência, base de cálculo, alíquota e valor do
'crédito de PIS/Pasep, referente às operações de transporte contratadas.

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_Processa_Registro_D101

    '******** registro D101 *****************
    objEFD.objSpedFiscal.sTipo = "D101"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D101", sChave)
    
    'Por enquanto só vai ter um indicador por NF
    'Já que o conhecimento de transporte só tem uma NF Original
    For Each objItem In objDoc.colItens
        Exit For
    Next
    
    lContD101 = lContD101 + 1
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D101")
    Call objEFD.Reg_Campo_add(2, "IND_NAT_FRT", objDoc.sIndNatFrtPis)
    Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objDoc.dValorMercadoria))
    Call objEFD.Reg_Campo_add(4, "CST_PIS", objItem.sPISCST)
    Call objEFD.Reg_Campo_add(5, "NAT_BC_CRED", SFF_Texto(objItem.sNatBCCred))
    Call objEFD.Reg_Campo_add(6, "VL_BC_PIS", SFF_Num(objDoc.dValorBasePIS))
    Call objEFD.Reg_Campo_add(8, "ALIQ_PIS", SFF_Perc(objItem.dPISAliquota))
    Call objEFD.Reg_Campo_add(8, "VL_PIS", SFF_Num(objDoc.dValorPIS))
    Call objEFD.Reg_Campo_add(9, "COD_CTA", "")
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    Processa_Registro_D101 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D101:

   Processa_Registro_D101 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211072)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D101", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D105(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD105 As Long) As Long
'REGISTRO D105: COMPLEMENTO DO DOCUMENTO DE TRANSPORTE (Códigos 07, 08, 8B, 09, 10, 11, 26, 27 e 57) - COFINS
'Serão escrituradas neste registro as informações referentes à incidência, base de cálculo, alíquota e valor do
'crédito de Cofins, referente às operações de transporte contratadas.

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_Processa_Registro_D105

    '******** registro D105 *****************
    objEFD.objSpedFiscal.sTipo = "D105"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D105", sChave)
    
    'Por enquanto só vai ter um indicador por NF
    'Já que o conhecimento de transporte só tem uma NF Original
    For Each objItem In objDoc.colItens
        Exit For
    Next
    
    lContD105 = lContD105 + 1
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D105")
    Call objEFD.Reg_Campo_add(2, "IND_NAT_FRT", objDoc.sIndNatFrtCofins)
    Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objDoc.dValorMercadoria))
    Call objEFD.Reg_Campo_add(4, "CST_COFINS", objItem.sPISCST)
    Call objEFD.Reg_Campo_add(5, "NAT_BC_CRED", SFF_Texto(objItem.sNatBCCred))
    Call objEFD.Reg_Campo_add(6, "VL_BC_COFINS", SFF_Num(objDoc.dValorBaseCofins))
    Call objEFD.Reg_Campo_add(8, "ALIQ_COFINS", SFF_Perc(objItem.dCOFINSAliquota))
    Call objEFD.Reg_Campo_add(8, "VL_COFINS", SFF_Num(objDoc.dValorCofins))
    Call objEFD.Reg_Campo_add(9, "COD_CTA", "")
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    Processa_Registro_D105 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D105:

   Processa_Registro_D105 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211073)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D105", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D111(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD111 As Long) As Long
'REGISTRO D111: PROCESSO REFERENCIADO

On Error GoTo Erro_Processa_Registro_D111

    '******** registro D111 *****************
    objEFD.objSpedFiscal.sTipo = "D111"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_D111 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D111:

   Processa_Registro_D111 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211074)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_D500(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD501 As Long, lContD505 As Long, lContD509 As Long) As Long
'REGISTRO D500: NOTA FISCAL DE SERVIÇO DE COMUNICAÇÃO (CÓDIGO 21) E NOTA FISCAL DE SERVIÇO DE TELECOMUNICAÇÃO
'(CÓDIGO 22) - DOCUMENTOS DE AQUISIÇÃO COM DIREITO A CRÉDITO
'Neste registro deverá a pessoa jurídica informar as operações referentes à contratação de serviços de comunicação
'ou de telecomunicação que, em função da natureza do serviço e da atividade econômica desenvolvida pela pessoa
'jurídica, permita a apuração de créditos de PIS/Pasep e de Cofins, na forma da legislação tributária.

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String

On Error GoTo Erro_Processa_Registro_D500

    '******** registro D500 *****************
    objEFD.objSpedFiscal.sTipo = "D500"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D500", sChave)
    
    If objDoc.lCliente > 0 Then
        sCodParticipante = "C" & Format(objDoc.lCliente, "0000000000") & Format(objDoc.iFilialCli, "00000")
    Else
        sCodParticipante = "F" & Format(objDoc.lFornecedor, "0000000000") & Format(objDoc.iFilialForn, "00000")
    End If
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D500")
    Call objEFD.Reg_Campo_add(2, "IND_OPER", objDoc.sOperacao)
    Call objEFD.Reg_Campo_add(3, "IND_EMIT", objDoc.sEmitente)
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
            Call objEFD.Reg_Campo_add(4, "COD_PART", "")
        
        Case Else
            Call objEFD.Reg_Campo_add(4, "COD_PART", sCodParticipante)
            
    End Select
    
    Call objEFD.Reg_Campo_add(5, "COD_MOD", objDoc.sModelo)
    Call objEFD.Reg_Campo_add(6, "COD_SIT", Format(objDoc.iSituacao, "00"))
    Call objEFD.Reg_Campo_add(7, "SER", SFF_Texto(objDoc.sSerie))
    Call objEFD.Reg_Campo_add(8, "SUB", "")
    Call objEFD.Reg_Campo_add(9, "NUM_DOC", CStr(objDoc.lNumDocumento))
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
            Call objEFD.Reg_Campo_add(10, "DT_DOC", "")
            Call objEFD.Reg_Campo_add(11, "DT_A_P", "")
            Call objEFD.Reg_Campo_add(12, "VL_DOC", "")
            Call objEFD.Reg_Campo_add(13, "VL_DESC", "")
            Call objEFD.Reg_Campo_add(14, "VL_SERV", "")
            Call objEFD.Reg_Campo_add(15, "VL_SERV_NT", "")
            Call objEFD.Reg_Campo_add(16, "VL_TERC", "")
            Call objEFD.Reg_Campo_add(17, "VL_DA", "")
            Call objEFD.Reg_Campo_add(18, "VL_BC_ICMS", "")
            Call objEFD.Reg_Campo_add(19, "VL_ICMS", "")
            Call objEFD.Reg_Campo_add(20, "COD_INF", "")
            Call objEFD.Reg_Campo_add(21, "VL_PIS", "")
            Call objEFD.Reg_Campo_add(22, "VL_COFINS", "")
        
        Case Else
        
            Call objEFD.Reg_Campo_add(10, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
            Call objEFD.Reg_Campo_add(11, "DT_A_P", SFF_Data(objDoc.dtDataES))
            Call objEFD.Reg_Campo_add(12, "VL_DOC", SFF_Num(objDoc.dValorTotal))
            Call objEFD.Reg_Campo_add(13, "VL_DESC", SFF_Num(objDoc.dValorDesconto))
            Call objEFD.Reg_Campo_add(14, "VL_SERV", SFF_Num(objDoc.dValorServico))
            Call objEFD.Reg_Campo_add(15, "VL_SERV_NT", SFF_Num(objDoc.dValorNaoTributado))
            Call objEFD.Reg_Campo_add(16, "VL_TERC", "")
            Call objEFD.Reg_Campo_add(17, "VL_DA", SFF_Num(objDoc.dValorDespesas))
            Call objEFD.Reg_Campo_add(18, "VL_BC_ICMS", SFF_Num(objDoc.dValorBaseICMS))
            Call objEFD.Reg_Campo_add(19, "VL_ICMS", SFF_Num(objDoc.dValorICMS))
            If objDoc.lCodMensagem > 0 Then
                Call objEFD.Reg_Campo_add(20, "COD_INF", Format(objDoc.lCodMensagem, "000000"))
            Else
                Call objEFD.Reg_Campo_add(20, "COD_INF", "")
            End If
            Call objEFD.Reg_Campo_add(21, "VL_PIS", SFF_Num(objDoc.dValorPIS))
            Call objEFD.Reg_Campo_add(22, "VL_COFINS", SFF_Num(objDoc.dValorCofins))
    
    End Select
    
    Call objEFD.Finaliza_Reg
    
    lSeq = objEFD.objSpedFiscal.lSeq

    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Select Case objDoc.iSituacao
    
        Case 2, 4, 5
    
        Case Else
            
            lErro = Processa_Registro_D501(objEFD, objDoc, lSeq, lContD501)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            lErro = Processa_Registro_D505(objEFD, objDoc, lSeq, lContD505)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Processa_Registro_D509(objEFD, objDoc, lSeq, lContD509)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End Select
    
    Processa_Registro_D500 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D500:

   Processa_Registro_D500 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211075)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D500", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D501(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD501 As Long) As Long
'REGISTRO D501: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 21 e 22) - PIS/PASEP

Dim lErro As Long
Dim sCodParticipante As String
Dim sChave As String
Dim objItem As ClassSpedDocFiscaisItens
Dim colAux As New Collection, bAchou As Boolean
Dim objItemAux As ClassSpedDocFiscaisItens


On Error GoTo Erro_Processa_Registro_D501

    '******** registro D501 *****************
    objEFD.objSpedFiscal.sTipo = "D501"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D501", sChave)
    
    For Each objItem In objDoc.colItens
        bAchou = False
        For Each objItemAux In colAux
            If objItemAux.sPISCST = objItem.sPISCST Then
                objItemAux.dValorBasePIS = objItemAux.dValorBasePIS + objItem.dValorBasePIS
                objItemAux.dValorPIS = objItemAux.dValorPIS + objItem.dValorPIS
                objItemAux.dPrecoTotal = objItemAux.dPrecoTotal + objItem.dPrecoTotal
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemAux = New ClassSpedDocFiscaisItens
            objItemAux.sPISCST = objItem.sPISCST
            objItemAux.dValorBasePIS = objItem.dValorBasePIS
            objItemAux.dValorPIS = objItem.dValorPIS
            objItemAux.dPrecoTotal = objItem.dPrecoTotal
            objItemAux.dPISAliquota = objItem.dPISAliquota
            objItemAux.sNatBCCred = objItem.sNatBCCred
            colAux.Add objItemAux
        End If
    Next
    
    For Each objItemAux In colAux
        
        lContD501 = lContD501 + 1
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "D501")
        Call objEFD.Reg_Campo_add(2, "CST_PIS", objItemAux.sPISCST)
        Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objItemAux.dPrecoTotal))
        Call objEFD.Reg_Campo_add(4, "NAT_BC_CRED", SFF_Texto(objItemAux.sNatBCCred))
        Call objEFD.Reg_Campo_add(5, "VL_BC_PIS", SFF_Num(objItemAux.dValorBasePIS))
        Call objEFD.Reg_Campo_add(6, "ALIQ_PIS", SFF_Perc(objItemAux.dPISAliquota))
        Call objEFD.Reg_Campo_add(7, "VL_PIS", SFF_Num(objItemAux.dValorPIS))
        Call objEFD.Reg_Campo_add(8, "COD_CTA", "")
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
       
    Processa_Registro_D501 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D501:

   Processa_Registro_D501 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211076)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D501", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D505(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD505 As Long) As Long
'REGISTRO D505: COMPLEMENTO DA OPERAÇÃO (CÓDIGOS 21 e 22) - COFINS

Dim lErro As Long
Dim sCodParticipante As String
Dim sChave As String
Dim objItem As ClassSpedDocFiscaisItens
Dim colAux As New Collection, bAchou As Boolean
Dim objItemAux As ClassSpedDocFiscaisItens

On Error GoTo Erro_Processa_Registro_D505

    '******** registro D505 *****************
    objEFD.objSpedFiscal.sTipo = "D505"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    sChave = "TipoDoc: " & CStr(objDoc.iTipoDoc) & " ID: " & CStr(objDoc.lNumIntDoc) & " Doc: " & objDoc.sSerie & " " & CStr(objDoc.lNumDocumento)
    Call objEFD.objTela.SPED_Status("Processando Registro D505", sChave)
    
    For Each objItem In objDoc.colItens
        bAchou = False
        For Each objItemAux In colAux
            If objItemAux.sCofinsCST = objItem.sCofinsCST Then
                objItemAux.dValorBaseCofins = objItemAux.dValorBaseCofins + objItem.dValorBaseCofins
                objItemAux.dValorCofins = objItemAux.dValorCofins + objItem.dValorCofins
                objItemAux.dPrecoTotal = objItemAux.dPrecoTotal + objItem.dPrecoTotal
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemAux = New ClassSpedDocFiscaisItens
            objItemAux.sCofinsCST = objItem.sCofinsCST
            objItemAux.dValorBaseCofins = objItem.dValorBaseCofins
            objItemAux.dValorCofins = objItem.dValorCofins
            objItemAux.dPrecoTotal = objItem.dPrecoTotal
            objItemAux.dCOFINSAliquota = objItem.dCOFINSAliquota
            objItemAux.sNatBCCred = objItem.sNatBCCred
            colAux.Add objItemAux
        End If
    Next
    
    For Each objItemAux In colAux
        
        lContD505 = lContD505 + 1
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "D505")
        Call objEFD.Reg_Campo_add(2, "CST_COFINS", objItemAux.sCofinsCST)
        Call objEFD.Reg_Campo_add(3, "VL_ITEM", SFF_Num(objItemAux.dPrecoTotal))
        Call objEFD.Reg_Campo_add(4, "NAT_BC_CRED", SFF_Texto(objItemAux.sNatBCCred))
        Call objEFD.Reg_Campo_add(5, "VL_BC_COFINS", SFF_Num(objItemAux.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(6, "ALIQ_COFINS", SFF_Perc(objItemAux.dCOFINSAliquota))
        Call objEFD.Reg_Campo_add(7, "VL_COFINS", SFF_Num(objItemAux.dValorCofins))
        Call objEFD.Reg_Campo_add(8, "COD_CTA", "")
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
       
    Processa_Registro_D505 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D505:

   Processa_Registro_D505 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211077)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_D505", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_D509(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long, lContD509 As Long) As Long
'REGISTRO D509: PROCESSO REFERENCIADO

On Error GoTo Erro_Processa_Registro_D509

    '******** registro D509 *****************
    objEFD.objSpedFiscal.sTipo = "D509"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    Processa_Registro_D509 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D509:

   Processa_Registro_D509 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211078)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_D990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO D990: ENCERRAMENTO DO BLOCO D
'Este registro destina-se a identificar o encerramento do bloco D e informar a quantidade de linhas (registros) existentes no
'bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_D990

    '******************* registro D990 **********************************
    objEFD.objSpedFiscal.sTipo = "D990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "D" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "D990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_D", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_D990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_D990:

   Processa_Registro_D990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201278)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_F001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBloco As Boolean) As Long
'REGISTRO F001: ABERTURA DO BLOCO F
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_F001

    '******** registro F001 *****************
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE  NumIntSped = ? AND Bloco = ?", lContador, objEFD.lNumIntSped, "F")
    If lErro <> AD_SQL_SUCESSO Then gError 201279
 
    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201280
        
    objEFD.objSpedFiscal.sTipo = "F001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBloco = True
    Else
        iTemReg = 1
        bAbreBloco = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "F001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_F001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_F001:

   Processa_Registro_F001 = gErr

   Select Case gErr
   
        Case 201279, 201280
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201281)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_F010(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO F010: IDENTIFICAÇÃO DO ESTABELECIMENTO
'Este registro tem o objetivo de identificar o estabelecimento da pessoa jurídica a que se referem as operações e
'documentos fiscais informados neste bloco. Só devem ser escriturados no Registro F010 os estabelecimentos da pessoa
'jurídica que efetivamente tenham realizado operações passíveis de escrituração neste bloco.
'O estabelecimento que não realizou operações passíveis de registro no bloco F, no período da escrituração, não deve ser
'identificado no Registro F010.
'Para cada estabelecimento cadastrado em “F010”, deve ser informado nos registros de nível inferior (Registros Filhos) as
'operações próprias do bloco, que tenham sido praticadas no período da escrituração.
Dim lErro As Long, iIndice As Integer
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 0) As Long
Dim sChave As String

On Error GoTo Erro_Processa_Registro_F010

    '******************* registro F010 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    tDoc.sCgc = String(STRING_CGC, 0)

    lErro = Comando_Executar(alComando(0), "SELECT DISTINCT CGC FROM SpedDocFiscais WHERE NumIntSped = ? ORDER BY CGC ", tDoc.sCgc, objEFD.lNumIntSped)
    If lErro <> AD_SQL_SUCESSO Then gError 209481
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209482
   
    Do While lErro = AD_SQL_SUCESSO

        sChave = "CNPJ\CPF: " & tDoc.sCgc
        Call objEFD.objTela.SPED_Status("Processando Registro F010", sChave)
        
        objEFD.objSpedFiscal.sTipo = "F010"
        objEFD.objSpedFiscal.lSeqPai = 0
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "F010")
        Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(tDoc.sCgc, 14))
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_F500(objEFD, objEFD.objSpedFiscal.lSeq, tDoc.sCgc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Processa_Registro_F525(objEFD, objEFD.objSpedFiscal.lSeq, tDoc.sCgc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209483

    Loop
        
    Processa_Registro_F010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_F010:

   Processa_Registro_F010 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case 209481 To 209483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209484)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_F010", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_F500(ByVal objEFD As ClassEFDPisCofinsSel, ByVal lSeqPai As Long, ByVal sCgc As String) As Long
'REGISTRO F500: CONSOLIDAÇÃO DAS OPERAÇÕES DA PESSOA JURÍDICA
'SUBMETIDA AO REGIME DE TRIBUTAÇÃO COM BASE NO LUCRO PRESUMIDO –
'INCIDÊNCIA DO PIS/PASEP E DA COFINS PELO REGIME DE CAIXA
'Registro especifico para a pessoa jurídica submetida ao regime de apuração com base no lucro presumido, optante pela
'apuração da contribuição para o PIS/Pasep e da Cofins pelo regime de caixa, conforme previsto no art. 20 da Medida
'Provisória nº 2.158-35, de 2001.
'Este registro tem por objetivo representar a escrituração e tratamento fiscal das receitas recebidas no período, segmentado
'por Código de Situação Tributária - CST, do PIS/Pasep e da Cofins. O total das receitas consolidadas por CST nos
'registros F500, devem corresponder ao total das receitas relacionadas nos registros F525.
'Os campos de CFOP, COD_CTA e INFO_COMPL podem ser utilizados pela pessoa jurídica para realizar o
'detalhamento da receita recebida por código de operação, documentos, contas contábeis, itens, clientes, etc.
Dim lErro As Long, iIndice As Integer
Dim tItem As typeSpedDocFiscaisItens
Dim alComando(0 To 0) As Long
Dim sChave As String
Dim sSQL As String, sModelo As String
Dim sPISCSTAnt As String, sCofinsCSTAnt As String
Dim dPISAliquotaAnt As Double, dCofinsAliquotaAnt As Double
Dim colItens As New Collection, colDocs As New Collection

On Error GoTo Erro_Processa_Registro_F500

    '******************* registro F500 **********************************
    objEFD.objSpedFiscal.sTipo = "F500"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    tItem.sPISCST = String(STRING_MAXIMO, 0)
    tItem.sCofinsCST = String(STRING_MAXIMO, 0)
    sModelo = String(STRING_MAXIMO, 0)
    tItem.sCFOP = String(STRING_MAXIMO, 0)

    sSQL = "SELECT D.Modelo, I.CFOP, I.PISCST, I.PISAliquota, I.COFINSCST, I.CofinsAliquota, "
    sSQL = sSQL + " " + "SUM (I.ValorRecebido), SUM(I.ValorBasePIS), SUM(I.ValorPis), SUM(I.ValorBaseCofins), SUM(I.ValorCofins) "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Bloco = 'F' AND D.Registro = 'F500' AND I.NumIntSped = ? AND D.CGC = ?"
    sSQL = sSQL + " " + "GROUP BY D.Modelo, I.CFOP, I.PISCST, I.PISAliquota, I.COFINSCST, I.CofinsAliquota"
    sSQL = sSQL + " " + "ORDER BY D.Modelo, I.CFOP, I.PISCST, I.PISAliquota, I.COFINSCST, I.CofinsAliquota"

    lErro = Comando_Executar(alComando(0), sSQL, sModelo, tItem.sCFOP, tItem.sPISCST, tItem.dPISAliquota, tItem.sCofinsCST, tItem.dCOFINSAliquota, tItem.dValorRecebido, tItem.dValorBasePIS, tItem.dValorPIS, tItem.dValorBaseCofins, tItem.dValorCofins, objEFD.lNumIntSped, sCgc)
    If lErro <> SUCESSO Then gError 211975

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211976

    Do While lErro = AD_SQL_SUCESSO
        
        sChave = "PIS CST: " & tItem.sPISCST & " COFINS CST: " & tItem.sCofinsCST
        Call objEFD.objTela.SPED_Status("Processando Registro F500", sChave)

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "F500")
        Call objEFD.Reg_Campo_add(2, "VL_REC_CAIXA", SFF_Num(tItem.dValorRecebido))
        Call objEFD.Reg_Campo_add(3, "CST_PIS", tItem.sPISCST)
        Call objEFD.Reg_Campo_add(4, "VL_DESC_PIS", "")
        Call objEFD.Reg_Campo_add(5, "VL_BC_PIS", SFF_Num(tItem.dValorBasePIS))
        Call objEFD.Reg_Campo_add(6, "ALIQ_PIS", SFF_Perc(tItem.dPISAliquota, 4))
        Call objEFD.Reg_Campo_add(7, "VL_PIS", SFF_Num(tItem.dValorPIS))
        Call objEFD.Reg_Campo_add(8, "CST_COFINS", tItem.sCofinsCST)
        Call objEFD.Reg_Campo_add(9, "VL_DESC_COFINS", "")
        Call objEFD.Reg_Campo_add(10, "VL_BC_COFINS", SFF_Num(tItem.dValorBaseCofins))
        Call objEFD.Reg_Campo_add(11, "ALIQ_COFINS", SFF_Perc(tItem.dCOFINSAliquota, 4))
        Call objEFD.Reg_Campo_add(12, "VL_COFINS", SFF_Num(tItem.dValorCofins))
        Call objEFD.Reg_Campo_add(13, "COD_MOD", sModelo)
        Call objEFD.Reg_Campo_add(14, "CFOP", tItem.sCFOP)
        Call objEFD.Reg_Campo_add(15, "COD_CTA", "")
        Call objEFD.Reg_Campo_add(16, "INFO_COMPL", "")
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211977

    Loop

    Processa_Registro_F500 = SUCESSO

    Exit Function

Erro_Processa_Registro_F500:

   Processa_Registro_F500 = gErr

   Select Case gErr
   
        Case 211975 To 211977
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211978)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_F500", sChave)

    Exit Function

End Function

Private Function Processa_Registro_F525(ByVal objEFD As ClassEFDPisCofinsSel, ByVal lSeqPai As Long, ByVal sCgc As String) As Long
'REGISTRO F525: COMPOSIÇÃO DA RECEITA ESCRITURADA NO PERÍODO –
'DETALHAMENTO DA RECEITA RECEBIDA PELO REGIME DE CAIXA
'Registro obrigatório para a pessoa jurídica submetida ao regime de tributação com base no lucro presumido, optante pela
'apuração das contribuições sociais pelo regime de caixa. Tem por objetivo relacionar a composição de todas as receitas
'recebidas pela pessoa jurídica no período da escrituração, sujeitas ou não ao pagamento da contribuição social.
'O total das receitas relacionadas nos registros F525 devem corresponder ao total das receitas recebidas, relacionadas nos
'registros F500.
'Atenção: Este registro é de escrituração opcional até o período de apuração referente a março de 2013. A partir de abril
'de 2013 o registro “F525” passa a ser de escrituração obrigatória.
Dim lErro As Long, iIndice As Integer
Dim tDoc As typeSpedDocFiscais, tItem As typeSpedDocFiscaisItens
Dim alComando(0 To 0) As Long
Dim sCGCAnt As String, sChave As String
Dim sSQL As String
Dim sPISCSTAnt As String, sCofinsCSTAnt As String
Dim dPISAliquotaAnt As Double, dCofinsAliquotaAnt As Double
Dim colItens As New Collection, colDocs As New Collection

On Error GoTo Erro_Processa_Registro_F525

    '******************* registro F525 **********************************
    objEFD.objSpedFiscal.sTipo = "F525"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    tItem.sPISCST = String(STRING_MAXIMO, 0)
    tItem.sCofinsCST = String(STRING_MAXIMO, 0)

    sSQL = "SELECT D.NumDocumento, I.PISCST, I.COFINSCST,"
    sSQL = sSQL + " " + "SUM (I.ValorRecebido), SUM(I.ValorBasePIS), SUM(I.ValorPis), SUM(I.ValorBaseCofins), SUM(I.ValorCofins) "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Bloco = 'F' AND D.Registro = 'F500' AND I.NumIntSped = ? AND D.CGC = ? "
    sSQL = sSQL + " " + "GROUP BY D.NumDocumento, I.PISCST, I.COFINSCST"
    sSQL = sSQL + " " + "ORDER BY D.NumDocumento, I.PISCST, I.COFINSCST"

    lErro = Comando_Executar(alComando(0), sSQL, tDoc.lNumDocumento, tItem.sPISCST, tItem.sCofinsCST, tItem.dValorRecebido, tItem.dValorBasePIS, tItem.dValorPIS, tItem.dValorBaseCofins, tItem.dValorCofins, objEFD.lNumIntSped, sCgc)
    If lErro <> SUCESSO Then gError 211979

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211980

    Do While lErro = AD_SQL_SUCESSO
    
        sChave = "Núm.Doc: " & CStr(tDoc.lNumDocumento)
        Call objEFD.objTela.SPED_Status("Processando Registro F525", sChave)

        objEFD.objSpedFiscal.sTipo = "F525"
        objEFD.objSpedFiscal.lSeqPai = lSeqPai

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "F525")
        Call objEFD.Reg_Campo_add(2, "VL_REC", SFF_Num(tItem.dValorRecebido))
        Call objEFD.Reg_Campo_add(3, "IND_REC", "04")
        Call objEFD.Reg_Campo_add(4, "CNPJ_CPF", "")
        Call objEFD.Reg_Campo_add(5, "NUM_DOC", CStr(tDoc.lNumDocumento))
        Call objEFD.Reg_Campo_add(6, "COD_ITEM", "")
        Call objEFD.Reg_Campo_add(7, "VL_REC_DET", SFF_Num(tItem.dValorRecebido))
        Call objEFD.Reg_Campo_add(8, "CST_PIS", tItem.sPISCST)
        Call objEFD.Reg_Campo_add(9, "CST_COFINS", tItem.sCofinsCST)
        Call objEFD.Reg_Campo_add(10, "INFO_COMPL", "")
        Call objEFD.Reg_Campo_add(11, "COD_CTA", "")
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211981

    Loop

    Processa_Registro_F525 = SUCESSO

    Exit Function

Erro_Processa_Registro_F525:

   Processa_Registro_F525 = gErr

   Select Case gErr
   
        Case 211979 To 211981
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211982)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_F525", sChave)

    Exit Function

End Function

Private Function Processa_Registro_F990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO F990: ENCERRAMENTO DO BLOCO F
'Este registro destina-se a identificar o encerramento do bloco F e informar a quantidade de linhas (registros) existentes no
'bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_F990

    '******************* registro F990 **********************************
    objEFD.objSpedFiscal.sTipo = "F990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "F" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "F990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_D", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_F990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_F990:

   Processa_Registro_F990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201282)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_M001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBloco As Boolean) As Long
'REGISTRO M001: ABERTURA DO BLOCO M
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_M001

    '******** registro M001 *****************
    'Abertura de Comando
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE  NumIntSped = ? AND Bloco = ?", lContador, objEFD.lNumIntSped, "M")
    If lErro <> AD_SQL_SUCESSO Then gError 201283

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201284
        
    objEFD.objSpedFiscal.sTipo = "M001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBloco = True
    Else
        iTemReg = 1
        bAbreBloco = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "M001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_M001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_M001:

   Processa_Registro_M001 = gErr

   Select Case gErr
   
        Case 201283, 201284
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201285)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_M200(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO M200: CONSOLIDAÇÃO DA CONTRIBUIÇÃO PARA O PIS/PASEP DO PERÍODO
'Neste registro serão consolidadas as contribuições sociais apuradas no período da escrituração, nos regimes nãocumulativo
'e cumulativo, bem como procedido ao desconto dos créditos não cumulativos apurados no próprio período,
'dos créditos apurados em períodos anteriores, dos valores retidos na fonte e de outras deduções previstas em Lei,
'demonstrando em seu final os valores devidos a recolher.
'ATENÇÃO: Os valores referentes às contribuições sociais escriturados nos Campos 02 e 09 do Registro M200 serão
'gerados com base:
'- Nos valores informados no arquivo elaborado pela própria pessoa jurídica e importado pelo Programa Validador
'e Assinador da EFD-Contribuições – PVA, os quais serão objeto de validação; ou
'- Nos valores das contribuições calculados pelo PVA no Registro M210 (Detalhamento da Contribuição para o
'PIS/Pasep no Período), no Campo 13 (VL_CONT_PER), através da funcionalidade “Gerar Apurações”,
'disponibilizada no PVA, com base nos registros de escrituração de receitas constantes nos Blocos “A”, “C”, “D”
'e “F”.
'A geração automática de apuração (funcionalidade “Gerar Apurações” (Ctrl+M)) o PVA apura, em relação ao Registro
'M200, apenas os valores dos campos de contribuições (Campos 02 e 09) e de créditos a descontar (Campos 03 e 04). Os
'campos representativos de retenções na fonte (Campos 06 e 10) e de outras deduções (07 e 11) não são recuperados na
'geração automática de apuração, devendo sempre ser informados pela própria pessoa jurídica no arquivo importado pelo
'PVA ou complementado pela edição (digitação no próprio PVA) do registro M200, dos respectivos valores de retenção
'na fonte escriturados nos registros F600, 1300 (PIS) ou 1700 (Cofins), e de deduções, escriturados no registro F700.
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 0) As Long, tDoc As typeSpedDocFiscais
Dim sChave As String, lSeqPai As Long, sSQL As String
Dim iCamposNovos As Integer

On Error GoTo Erro_Processa_Registro_M200

    '******************* registro M200 **********************************
    
    If objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = DESMARCADO Then
    'Para o regime cumulativo
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = objEFD.ProxComando
        Next
               
        objEFD.objSpedFiscal.sTipo = "M200"
        objEFD.objSpedFiscal.lSeqPai = 0
        sChave = ""
        Call objEFD.objTela.SPED_Status("Processando Registro M200", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "M200")
        Call objEFD.Reg_Campo_add(2, "VL_TOT_CONT_NC_PER", SFF_Num(0))
        Call objEFD.Reg_Campo_add(3, "VL_TOT_CRED_DESC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(4, "VL_TOT_CRED_DESC_ANT", SFF_Num(0))
        Call objEFD.Reg_Campo_add(5, "VL_TOT_CONT_NC_DEV", SFF_Num(0))
        Call objEFD.Reg_Campo_add(6, "VL_RET_NC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(7, "VL_OUT_DED_NC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(8, "VL_CONT_NC_REC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(9, "VL_TOT_CONT_CUM_PER", SFF_Num(0))
        Call objEFD.Reg_Campo_add(10, "VL_RET_CUM", SFF_Num(0))
        Call objEFD.Reg_Campo_add(11, "VL_OUT_DED_CUM", SFF_Num(0))
        Call objEFD.Reg_Campo_add(12, "VL_CONT_CUM_REC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(13, "VL_TOT_CONT_REC", SFF_Num(0))
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lSeqPai = objEFD.objSpedFiscal.lSeq
        
        objEFD.objSpedFiscal.sTipo = "M210"
        objEFD.objSpedFiscal.lSeqPai = lSeqPai
        sChave = ""
        Call objEFD.objTela.SPED_Status("Processando Registro M210", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "M210")
        Call objEFD.Reg_Campo_add(2, "COD_CONT", "51")
        Call objEFD.Reg_Campo_add(3, "VL_REC_BRT", SFF_Num(0))
        Call objEFD.Reg_Campo_add(4, "VL_BC_CONT", SFF_Num(0))
        
        iCamposNovos = 0
        
        If objEFD.iCodVersao >= 5 Then 'posterior à 01/01/2019
            Call objEFD.Reg_Campo_add(5, "VL_AJUS_ACRES_BC_PIS", SFF_Num(0))
            Call objEFD.Reg_Campo_add(6, "VL_AJUS_REDUC_BC_PIS", SFF_Num(0))
            Call objEFD.Reg_Campo_add(7, "VL_BC_CONT_AJUS", SFF_Num(0))
            iCamposNovos = 3
        End If
        
        Call objEFD.Reg_Campo_add(5 + iCamposNovos, "ALIQ_PIS", SFF_Perc(0.0065))
        Call objEFD.Reg_Campo_add(6 + iCamposNovos, "QUANT_BC_PIS", "")
        Call objEFD.Reg_Campo_add(7 + iCamposNovos, "ALIQ_PIS_QUANT", "")
        Call objEFD.Reg_Campo_add(8 + iCamposNovos, "VL_CONT_APUR", SFF_Num(0))
        Call objEFD.Reg_Campo_add(9 + iCamposNovos, "VL_AJUS_ACRES", SFF_Num(0))
        Call objEFD.Reg_Campo_add(10 + iCamposNovos, "VL_AJUS_REDUC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(11 + iCamposNovos, "VL_CONT_DIFER", "")
        Call objEFD.Reg_Campo_add(12 + iCamposNovos, "VL_CONT_DIFER_ANT", "")
        Call objEFD.Reg_Campo_add(13 + iCamposNovos, "VL_CONT_PER", SFF_Num(0))
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lSeqPai = objEFD.objSpedFiscal.lSeq
        
        
        sSQL = "SELECT NumDocumento, ValorPis, DataES FROM SpedDocFiscais "
        sSQL = sSQL + " " + "WHERE Bloco = 'M' AND Registro = 'M220' AND NumIntSped = ? "
        sSQL = sSQL + " " + "ORDER BY NumDocumento"
    
        lErro = Comando_Executar(alComando(0), sSQL, tDoc.lNumDocumento, tDoc.dValorPIS, tDoc.dtDataES, objEFD.lNumIntSped)
        If lErro <> SUCESSO Then gError 213068
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213069
    
        Do While lErro = AD_SQL_SUCESSO
            
            objEFD.objSpedFiscal.sTipo = "M220"
            objEFD.objSpedFiscal.lSeqPai = lSeqPai
            sChave = "NumDoc: " & CStr(tDoc.lNumDocumento)
            Call objEFD.objTela.SPED_Status("Processando Registro M220", sChave)
        
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "M220")
            Call objEFD.Reg_Campo_add(2, "IND_AJ", "0")
            Call objEFD.Reg_Campo_add(3, "VL_AJ", SFF_Num(tDoc.dValorPIS))
            Call objEFD.Reg_Campo_add(4, "COD_AJ", "06")
            Call objEFD.Reg_Campo_add(5, "NUM_DOC", SFF_Texto(tDoc.lNumDocumento))
            Call objEFD.Reg_Campo_add(6, "DESCR_AJ", SFF_Texto("DEVOLUÇÃO DE VENDA"))
            Call objEFD.Reg_Campo_add(7, "DT_REF", SFF_Data(tDoc.dtDataES))
            Call objEFD.Finaliza_Reg
    
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213070
    
        Loop
   
    End If

    Processa_Registro_M200 = SUCESSO

    Exit Function

Erro_Processa_Registro_M200:

   Processa_Registro_M200 = gErr

   Select Case gErr

        Case 213068 To 213070
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213071)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_M200", sChave)

    Exit Function

End Function

Private Function Processa_Registro_M600(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO M600: CONSOLIDAÇÃO DA CONTRIBUIÇÃO PARA A SEGURIDADE SOCIAL - COFINS DO PERÍODO
'Neste registro serão consolidadas as contribuições sociais apuradas no período da escrituração, nos regimes nãocumulativo
'e cumulativo, bem como procedido ao desconto dos créditos não cumulativos apurados no próprio período,
'dos créditos apurados em períodos anteriores, dos valores retidos na fonte e de outras deduções previstas em lei,
'demonstrando em seu final os valores devidos a recolher.
'ATENÇÃO: Os valores referentes às contribuições sociais escriturados nos Campos 02 e 09 do Registro M600 serão
'gerados com base:
'- Nos valores informados no arquivo elaborado pela própria pessoa jurídica e importado pelo Programa Validador
'e Assinador da EFD-Contribuições – PVA, os quais serão objeto de validação; ou
'- Nos valores das contribuições calculados pelo PVA no Registro M610 (Detalhamento da Cofins no Período), no
'Campo 13 (VL_CONT_PER), através da funcionalidade “Gerar Apurações”, disponibilizada no PVA, com base
'nos registros de escrituração de receitas constantes nos Blocos “A”, “C”, “D” e “F”.
'A geração automática de apuração (funcionalidade “Gerar Apurações” (Ctrl+M)) o PVA apura, em relação ao Registro
'M600, apenas os valores dos campos de contribuições (Campos 02 e 09) e de créditos a descontar (Campos 03 e 04). Os
'campos representativos de retenções na fonte (Campos 06 e 10) e de outras deduções (07 e 11) não serão recuperados na
'geração automática de apuração, devendo sempre ser informados pela própria pessoa jurídica no arquivo importado pelo
'PVA ou complementado pela edição (digitação no próprio PVA) no registro M600, dos respectivos valores de retenção
'na fonte escriturados nos registros F600, 1300 (PIS) ou 1700 (Cofins), e de deduções, escriturados no registro F700.
'Na funcionalidade de geração automática de apuração, os valores apurados e preenchidos pelo PVA para os Campos 02 e
'09 (contribuições apuradas) e para os Campos 03 e 04 (créditos descontados) irão sobrepor (substituir) os valores
'eventualmente existentes nos referidos campos, constantes na escrituração.
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 0) As Long, tDoc As typeSpedDocFiscais
Dim sChave As String, lSeqPai As Long, sSQL As String
Dim iCamposNovos As Integer

On Error GoTo Erro_Processa_Registro_M600

    '******************* registro M600 **********************************
    
    If objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = DESMARCADO Then
    'Para o regime cumulativo
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = objEFD.ProxComando
        Next
               
        objEFD.objSpedFiscal.sTipo = "M600"
        objEFD.objSpedFiscal.lSeqPai = 0
        sChave = ""
        Call objEFD.objTela.SPED_Status("Processando Registro M600", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "M600")
        Call objEFD.Reg_Campo_add(2, "VL_TOT_CONT_NC_PER", SFF_Num(0))
        Call objEFD.Reg_Campo_add(3, "VL_TOT_CRED_DESC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(4, "VL_TOT_CRED_DESC_ANT", SFF_Num(0))
        Call objEFD.Reg_Campo_add(5, "VL_TOT_CONT_NC_DEV", SFF_Num(0))
        Call objEFD.Reg_Campo_add(6, "VL_RET_NC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(7, "VL_OUT_DED_NC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(8, "VL_CONT_NC_REC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(9, "VL_TOT_CONT_CUM_PER", SFF_Num(0))
        Call objEFD.Reg_Campo_add(10, "VL_RET_CUM", SFF_Num(0))
        Call objEFD.Reg_Campo_add(11, "VL_OUT_DED_CUM", SFF_Num(0))
        Call objEFD.Reg_Campo_add(12, "VL_CONT_CUM_REC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(13, "VL_TOT_CONT_REC", SFF_Num(0))
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lSeqPai = objEFD.objSpedFiscal.lSeq
        
        objEFD.objSpedFiscal.sTipo = "M610"
        objEFD.objSpedFiscal.lSeqPai = lSeqPai
        sChave = ""
        Call objEFD.objTela.SPED_Status("Processando Registro M610", sChave)
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "M610")
        Call objEFD.Reg_Campo_add(2, "COD_CONT", "51")
        Call objEFD.Reg_Campo_add(3, "VL_REC_BRT", SFF_Num(0))
        Call objEFD.Reg_Campo_add(4, "VL_BC_CONT", SFF_Num(0))
        
        iCamposNovos = 0
        
        If objEFD.iCodVersao >= 5 Then 'posterior à 01/01/2019
            Call objEFD.Reg_Campo_add(5, "VL_AJUS_ACRES_BC_COFINS", SFF_Num(0))
            Call objEFD.Reg_Campo_add(6, "VL_AJUS_REDUC_BC_COFINS", SFF_Num(0))
            Call objEFD.Reg_Campo_add(7, "VL_BC_CONT_AJUS", SFF_Num(0))
            iCamposNovos = 3
        End If
        
        
        Call objEFD.Reg_Campo_add(5 + iCamposNovos, "ALIQ_COFINS", SFF_Perc(0.03))
        Call objEFD.Reg_Campo_add(6 + iCamposNovos, "QUANT_BC_COFINS", "")
        Call objEFD.Reg_Campo_add(7 + iCamposNovos, "ALIQ_COFINS_QUANT", "")
        Call objEFD.Reg_Campo_add(8 + iCamposNovos, "VL_CONT_APUR", SFF_Num(0))
        Call objEFD.Reg_Campo_add(9 + iCamposNovos, "VL_AJUS_ACRES", SFF_Num(0))
        Call objEFD.Reg_Campo_add(10 + iCamposNovos, "VL_AJUS_REDUC", SFF_Num(0))
        Call objEFD.Reg_Campo_add(11 + iCamposNovos, "VL_CONT_DIFER", "")
        Call objEFD.Reg_Campo_add(12 + iCamposNovos, "VL_CONT_DIFER_ANT", "")
        Call objEFD.Reg_Campo_add(13 + iCamposNovos, "VL_CONT_PER", SFF_Num(0))
        Call objEFD.Finaliza_Reg
    
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lSeqPai = objEFD.objSpedFiscal.lSeq
        
        sSQL = "SELECT NumDocumento, ValorCofins, DataES FROM SpedDocFiscais "
        sSQL = sSQL + " " + "WHERE Bloco = 'M' AND Registro = 'M220' AND NumIntSped = ? "
        sSQL = sSQL + " " + "ORDER BY NumDocumento"
    
        lErro = Comando_Executar(alComando(0), sSQL, tDoc.lNumDocumento, tDoc.dValorCofins, tDoc.dtDataES, objEFD.lNumIntSped)
        If lErro <> SUCESSO Then gError 213068
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213069
    
        Do While lErro = AD_SQL_SUCESSO
            
            objEFD.objSpedFiscal.sTipo = "M620"
            objEFD.objSpedFiscal.lSeqPai = lSeqPai
            sChave = "NumDoc: " & CStr(tDoc.lNumDocumento)
            Call objEFD.objTela.SPED_Status("Processando Registro M620", sChave)
        
            Call objEFD.Inicia_Reg
            Call objEFD.Reg_Campo_add(1, "REG", "M620")
            Call objEFD.Reg_Campo_add(2, "IND_AJ", "0")
            Call objEFD.Reg_Campo_add(3, "VL_AJ", SFF_Num(tDoc.dValorCofins))
            Call objEFD.Reg_Campo_add(4, "COD_AJ", "06")
            Call objEFD.Reg_Campo_add(5, "NUM_DOC", SFF_Texto(tDoc.lNumDocumento))
            Call objEFD.Reg_Campo_add(6, "DESCR_AJ", SFF_Texto("DEVOLUÇÃO DE VENDA"))
            Call objEFD.Reg_Campo_add(7, "DT_REF", SFF_Data(tDoc.dtDataES))
            Call objEFD.Finaliza_Reg
    
            lErro = Registro_Insere_BD(objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213070
    
        Loop
   
    End If

    Processa_Registro_M600 = SUCESSO

    Exit Function

Erro_Processa_Registro_M600:

   Processa_Registro_M600 = gErr

   Select Case gErr

        Case 213068 To 213070
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213071)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_M600", sChave)

    Exit Function

End Function

Private Function Processa_Registro_M990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO M990: ENCERRAMENTO DO BLOCO M
'Este registro destina-se a identificar o encerramento do bloco F e informar a quantidade de linhas (registros) existentes no
'bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_M990

    '******************* registro M990 **********************************
    objEFD.objSpedFiscal.sTipo = "M990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "M" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "M990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_D", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_M990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_M990:

   Processa_Registro_M990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201286)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_1001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBloco As Boolean) As Long
'REGISTRO 1001: ABERTURA DO BLOCO 1
Dim lErro As Long
Dim lContador1 As Long, lContador2 As Long
Dim iTemReg As Integer
Dim dtDataAux As Date, iMes As Integer, iAno As Integer

On Error GoTo Erro_Processa_Registro_1001

    '******** registro 1001 *****************

    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE NumIntSped = ? AND Bloco = ?", lContador1, objEFD.lNumIntSped, "1")
    If lErro <> AD_SQL_SUCESSO Then gError 201287

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201288
    
    dtDataAux = DateAdd("m", -1, objEFD.dtDataIni)
    iAno = Year(dtDataAux)
    iMes = Month(dtDataAux)
    
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Ano= ?  AND Mes= ?", lContador2, objEFD.iFilialEmpresa, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 201287

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201288
        
    objEFD.objSpedFiscal.sTipo = "1001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador1 + lContador2 > 0 Then
        iTemReg = 0
        bAbreBloco = True
    Else
        iTemReg = 1
        bAbreBloco = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "1001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_1001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_1001:

   Processa_Registro_1001 = gErr

   Select Case gErr
   
        Case 201287, 201288
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201289)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_1100(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 1100: CONTROLE DE CRÉDITOS FISCAIS – PIS/PASEP
'Este registro tem por objetivo realizar o controle de saldos de créditos fiscais de períodos anteriores ao da atual
'escrituração, bem como eventual saldo credor apurado no próprio período da escrituração. Ou seja, este registro serve
'para escriturar as disponibilidades de créditos:
'- Apurados em períodos anteriores ao da escrituração, demonstrados mês a mês, com saldos a utilizar no atual
'período da escrituração ou em períodos posteriores, mediante desconto, compensação ou ressarcimento;
'- Apurados no próprio período da escrituração, mas que não foi totalmente utilizado neste período, restando saldos
'a utilizar em períodos posteriores, mediante desconto, compensação ou ressarcimento.
'Atenção: Não precisam ser escriturados neste registro os créditos apurados no próprio período e que foram totalmente
'utilizados na atual escrituração, não restando assim saldos a utilizar em período posterior.
'O saldo de créditos deverá ser segregado por período de apuração, devendo, ainda, levar em consideração a sua origem e,
'no caso de créditos transferidos por sucessão, o CNPJ da pessoa jurídica cedente do crédito. A chave deste registro é
'formada pelo campo PER_APU_CRED, campo ORIG_CRED, campo CNPJ_SUC e campo COD_CRED.
'Conceitualmente, o crédito só se caracteriza como extemporâneo, quando se refere a período anterior ao da escrituração, e
'o mesmo não pode mais ser escriturado no correspondente período de apuração de sua constituição, via transmissão de
'Dacon retificador ou EFD-Contribuições retificadora.
'Desta forma, eventual crédito extemporâneo informado no campo 07 tem, necessariamente, que se referir a período de
'apuração (campo 02) anterior ao da atual escrituração.
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 0) As Long
Dim sChave As String, dtDataAux As Date, iMes As Integer, iAno As Integer
Dim sSQL As String, tApur As typeRegApuracaoPISCofins

On Error GoTo Erro_Processa_Registro_1100

    '******************* registro 1100 **********************************
    objEFD.objSpedFiscal.sTipo = "1100"
    objEFD.objSpedFiscal.lSeqPai = 0

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    dtDataAux = DateAdd("m", -1, objEFD.dtDataIni)
    iAno = Year(dtDataAux)
    iMes = Month(dtDataAux)
    
    'Alocação de espaço no buffer
    tApur.sCNPJCedCred = String(STRING_CGC, 0)
    tApur.sCodCred = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(alComando(0), "SELECT OrigCred, CNPJCedCred, CodCred, VlCredApu, VlCredExtApu, VlTotCredApu, VlCredDescPAAnt, VlCredPerPAAnt, VlCredDCompPAAnt, " & _
                "SdCredDispEFD, VlCredDescEFD, VlCredPerEFD, VlCredDCompEFD, VlCredTrans, VlCredOut, SdCredFim FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ? ", _
                tApur.iOrigCred, tApur.sCNPJCedCred, tApur.sCodCred, tApur.dVlCredApu, tApur.dVlCredExtApu, _
                tApur.dVlTotCredApu, tApur.dVlCredDescPAAnt, tApur.dVlCredPerPAAnt, tApur.dVlCredDCompPAAnt, tApur.dSdCredDispEFD, _
                tApur.dVlCredDescEFD, tApur.dVlCredPerEFD, tApur.dVlCredDCompEFD, tApur.dVlCredTrans, tApur.dVlCredOut, _
                tApur.dSdCredFim, objEFD.iFilialEmpresa, REG_APUR_PIS_COFINS_TIPO_PIS, iAno, iMes)
    If lErro <> SUCESSO Then gError 213064

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213065

    Do While lErro = AD_SQL_SUCESSO
        
        sChave = "OrigCred: " & CStr(tApur.iOrigCred) & " CNPJCedCred: " & tApur.sCNPJCedCred & " CodCred: " & tApur.sCodCred
        Call objEFD.objTela.SPED_Status("Processando Registro 1100", sChave)

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "1100")
        Call objEFD.Reg_Campo_add(2, "PER_APU_CRED", Format(iMes, "00") & CStr(iAno))
        Call objEFD.Reg_Campo_add(3, "ORIG_CRED", Format(tApur.iOrigCred, "00"))
        Call objEFD.Reg_Campo_add(4, "CNPJ_SUC", SFF_Texto(tApur.sCNPJCedCred, 14))
        Call objEFD.Reg_Campo_add(5, "COD_CRED", SFF_Texto(tApur.sCodCred))
        Call objEFD.Reg_Campo_add(6, "VL_CRED_APU", SFF_Num(tApur.dVlCredApu))
        Call objEFD.Reg_Campo_add(7, "VL_CRED_EXT_APU", SFF_Num(tApur.dVlCredExtApu))
        Call objEFD.Reg_Campo_add(8, "VL_TOT_CRED_APU", SFF_Num(tApur.dVlTotCredApu))
        Call objEFD.Reg_Campo_add(9, "VL_CRED_DESC_PA_ANT", SFF_Num(tApur.dVlCredDescPAAnt))
        Call objEFD.Reg_Campo_add(10, "VL_CRED_PER_PA_ANT", SFF_Num(tApur.dVlCredPerPAAnt))
        Call objEFD.Reg_Campo_add(11, "VL_CRED_DCOMP_PA_ANT", SFF_Num(tApur.dVlCredDCompPAAnt))
        Call objEFD.Reg_Campo_add(12, "SD_CRED_DISP_EFD", SFF_Num(tApur.dSdCredDispEFD))
        Call objEFD.Reg_Campo_add(13, "VL_CRED_DESC_EFD", SFF_Num(tApur.dVlCredDescEFD))
        Call objEFD.Reg_Campo_add(14, "VL_CRED_PER_EFD", SFF_Num(tApur.dVlCredPerEFD))
        Call objEFD.Reg_Campo_add(15, "VL_CRED_DCOMP_EFD", SFF_Num(tApur.dVlCredDCompEFD))
        Call objEFD.Reg_Campo_add(16, "VL_CRED_TRANS", SFF_Num(tApur.dVlCredTrans))
        Call objEFD.Reg_Campo_add(17, "VL_CRED_OUT", SFF_Num(tApur.dVlCredOut))
        Call objEFD.Reg_Campo_add(18, "SLD_CRED_FIM", SFF_Num(tApur.dSdCredFim))
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213066

    Loop

    Processa_Registro_1100 = SUCESSO

    Exit Function

Erro_Processa_Registro_1100:

   Processa_Registro_1100 = gErr

   Select Case gErr

        Case 213064 To 213066
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213067)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_1100", sChave)

    Exit Function

End Function

Private Function Processa_Registro_1500(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 1500: CONTROLE DE CRÉDITOS FISCAIS – COFINS
'Este registro tem por objetivo realizar o controle de saldos de créditos fiscais de períodos anteriores ao da atual
'escrituração, bem como eventual saldo credor apurado no próprio período da escrituração.
'O saldo de créditos deverá ser segregado por período de apuração, devendo, ainda, levar em consideração a sua origem e,
'no caso de créditos transferidos por sucessão, o CNPJ da pessoa jurídica cedente do crédito. A chave deste registro é
'formada pelo campo PER_APU_CRED, campo ORIG_CRED, campo CNPJ_SUC e campo COD_CRED. Este registro
'tem por objetivo realizar o controle de saldos de créditos fiscais de períodos anteriores ao da atual escrituração, bem
'como eventual saldo credor apurado no próprio período da escrituração.
'O saldo de créditos deverá ser segregado por período de apuração, devendo, ainda, levar em consideração a sua origem e,
'no caso de créditos transferidos por sucessão, o CNPJ da pessoa jurídica cedente do crédito. A chave deste registro é
'formada pelo campo PER_APU_CRED, campo ORIG_CRED, campo CNPJ_SUC e campo COD_CRED.
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 0) As Long
Dim sChave As String, dtDataAux As Date, iMes As Integer, iAno As Integer
Dim sSQL As String, tApur As typeRegApuracaoPISCofins

On Error GoTo Erro_Processa_Registro_1500

    '******************* registro 1500 **********************************
    objEFD.objSpedFiscal.sTipo = "1500"
    objEFD.objSpedFiscal.lSeqPai = 0

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    dtDataAux = DateAdd("m", -1, objEFD.dtDataIni)
    iAno = Year(dtDataAux)
    iMes = Month(dtDataAux)
    
    'Alocação de espaço no buffer
    tApur.sCNPJCedCred = String(STRING_CGC, 0)
    tApur.sCodCred = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(alComando(0), "SELECT OrigCred, CNPJCedCred, CodCred, VlCredApu, VlCredExtApu, VlTotCredApu, VlCredDescPAAnt, VlCredPerPAAnt, VlCredDCompPAAnt, " & _
                "SdCredDispEFD, VlCredDescEFD, VlCredPerEFD, VlCredDCompEFD, VlCredTrans, VlCredOut, SdCredFim FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ? ", _
                tApur.iOrigCred, tApur.sCNPJCedCred, tApur.sCodCred, tApur.dVlCredApu, tApur.dVlCredExtApu, _
                tApur.dVlTotCredApu, tApur.dVlCredDescPAAnt, tApur.dVlCredPerPAAnt, tApur.dVlCredDCompPAAnt, tApur.dSdCredDispEFD, _
                tApur.dVlCredDescEFD, tApur.dVlCredPerEFD, tApur.dVlCredDCompEFD, tApur.dVlCredTrans, tApur.dVlCredOut, _
                tApur.dSdCredFim, objEFD.iFilialEmpresa, REG_APUR_PIS_COFINS_TIPO_COFINS, iAno, iMes)
    If lErro <> SUCESSO Then gError 213068

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213069

    Do While lErro = AD_SQL_SUCESSO
        
        sChave = "OrigCred: " & CStr(tApur.iOrigCred) & " CNPJCedCred: " & tApur.sCNPJCedCred & " CodCred: " & tApur.sCodCred
        Call objEFD.objTela.SPED_Status("Processando Registro 1500", sChave)

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "1500")
        Call objEFD.Reg_Campo_add(2, "PER_APU_CRED", Format(iMes, "00") & CStr(iAno))
        Call objEFD.Reg_Campo_add(3, "ORIG_CRED", Format(tApur.iOrigCred, "00"))
        Call objEFD.Reg_Campo_add(4, "CNPJ_SUC", SFF_Texto(tApur.sCNPJCedCred, 14))
        Call objEFD.Reg_Campo_add(5, "COD_CRED", SFF_Texto(tApur.sCodCred))
        Call objEFD.Reg_Campo_add(6, "VL_CRED_APU", SFF_Num(tApur.dVlCredApu))
        Call objEFD.Reg_Campo_add(7, "VL_CRED_EXT_APU", SFF_Num(tApur.dVlCredExtApu))
        Call objEFD.Reg_Campo_add(8, "VL_TOT_CRED_APU", SFF_Num(tApur.dVlTotCredApu))
        Call objEFD.Reg_Campo_add(9, "VL_CRED_DESC_PA_ANT", SFF_Num(tApur.dVlCredDescPAAnt))
        Call objEFD.Reg_Campo_add(10, "VL_CRED_PER_PA_ANT", SFF_Num(tApur.dVlCredPerPAAnt))
        Call objEFD.Reg_Campo_add(11, "VL_CRED_DCOMP_PA_ANT", SFF_Num(tApur.dVlCredDCompPAAnt))
        Call objEFD.Reg_Campo_add(12, "SD_CRED_DISP_EFD", SFF_Num(tApur.dSdCredDispEFD))
        Call objEFD.Reg_Campo_add(13, "VL_CRED_DESC_EFD", SFF_Num(tApur.dVlCredDescEFD))
        Call objEFD.Reg_Campo_add(14, "VL_CRED_PER_EFD", SFF_Num(tApur.dVlCredPerEFD))
        Call objEFD.Reg_Campo_add(15, "VL_CRED_DCOMP_EFD", SFF_Num(tApur.dVlCredDCompEFD))
        Call objEFD.Reg_Campo_add(16, "VL_CRED_TRANS", SFF_Num(tApur.dVlCredTrans))
        Call objEFD.Reg_Campo_add(17, "VL_CRED_OUT", SFF_Num(tApur.dVlCredOut))
        Call objEFD.Reg_Campo_add(18, "SLD_CRED_FIM", SFF_Num(tApur.dSdCredFim))
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213070

    Loop

    Processa_Registro_1500 = SUCESSO

    Exit Function

Erro_Processa_Registro_1500:

   Processa_Registro_1500 = gErr

   Select Case gErr

        Case 213068 To 213070
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213071)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_1500", sChave)

    Exit Function

End Function

Private Function Processa_Registro_1900(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 1900: CONSOLIDAÇÃO DOS DOCUMENTOS EMITIDOS NO PERÍODO POR
'PESSOA JURÍDICA SUBMETIDA AO REGIME DE TRIBUTAÇÃO COM BASE NO
'LUCRO PRESUMIDO – REGIME DE CAIXA OU DE COMPETÊNCIA
'Registro para a pessoa jurídica sujeita à tributação com base no lucro presumido, que procedeu à escrituração de suas
'receitas de forma consolidada, pelo regime de caixa (registro “F500” ou “F510”) ou de competência (registro “F550” ou
'“F560”), informar o valor consolidado dos documentos fiscais e demais documentos, emitidos no período da escrituração,
'representativos de receitas da venda de bens e serviços efetuada no período, independente de sua realização (recebimento)
'ou não.
'Atenção: Este registro é de escrituração opcional até o período de apuração referente a março de 2013. A partir de abril
'de 2013 o registro “1900” passa a ser de escrituração obrigatória.
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 0) As Long
Dim sChave As String
Dim sSQL As String, sModelo As String, sCFOP As String
Dim sCgc As String, dValor As Double, lQtdeDocs As Long

On Error GoTo Erro_Processa_Registro_1900

    '******************* registro 1900 **********************************
    objEFD.objSpedFiscal.sTipo = "1900"
    objEFD.objSpedFiscal.lSeqPai = 0

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    sCgc = String(STRING_MAXIMO, 0)
    sModelo = String(STRING_MAXIMO, 0)
    sCFOP = String(STRING_MAXIMO, 0)

    sSQL = "SELECT D.CGC, D.Modelo, I.CFOP, SUM (I.ValorNoCR) "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Bloco = '1' AND D.Registro = '1900' AND I.NumIntSped = ? "
    sSQL = sSQL + " " + "GROUP BY D.CGC, D.Modelo, I.CFOP"
    sSQL = sSQL + " " + "HAVING SUM(I.ValorNoCR) > 0.001"
    sSQL = sSQL + " " + "ORDER BY D.CGC, D.Modelo, I.CFOP"

    lErro = Comando_Executar(alComando(0), sSQL, sCgc, sModelo, sCFOP, dValor, objEFD.lNumIntSped)
    If lErro <> SUCESSO Then gError 211970

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211971

    Do While lErro = AD_SQL_SUCESSO
        
        sChave = "CGC: " & sCgc & " Modelo: " & sModelo & " Valor: R$ " & Format(dValor, "STANDARD")
        Call objEFD.objTela.SPED_Status("Processando Registro 1900", sChave)

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "1900")
        Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(sCgc, 14))
        Call objEFD.Reg_Campo_add(3, "COD_MOD", sModelo)
        Call objEFD.Reg_Campo_add(4, "SER", "")
        Call objEFD.Reg_Campo_add(5, "SUB_SER", "")
        Call objEFD.Reg_Campo_add(6, "COD_SIT", "00")
        Call objEFD.Reg_Campo_add(7, "VL_TOT_REC", SFF_Num(dValor))
        Call objEFD.Reg_Campo_add(8, "QUANT_DOC", "")
        Call objEFD.Reg_Campo_add(9, "CST_PIS", "")
        Call objEFD.Reg_Campo_add(10, "CST_COFINS", "")
        Call objEFD.Reg_Campo_add(11, "CFOP", sCFOP)
        Call objEFD.Reg_Campo_add(12, "INF_COMPL", "")
        Call objEFD.Reg_Campo_add(13, "COD_CTA", "")
        Call objEFD.Finaliza_Reg

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211972

    Loop

    Processa_Registro_1900 = SUCESSO

    Exit Function

Erro_Processa_Registro_1900:

   Processa_Registro_1900 = gErr

   Select Case gErr

        Case 211970 To 211972
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211973)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_1900", sChave)

    Exit Function

End Function

Private Function Processa_Registro_1990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO 1990: ENCERRAMENTO DO BLOCO 1
'Este registro destina-se a identificar o encerramento do bloco 1 e informar a quantidade de linhas (registros) existentes no
'bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_1990

    '******************* registro 1990 **********************************
    objEFD.objSpedFiscal.sTipo = "1990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1

    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "1" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "1990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_D", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_1990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_1990:

   Processa_Registro_1990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201290)
    
    End Select
    
    Exit Function
    
End Function

Private Function Processa_Registro_P001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBlocoP As Boolean) As Long
'REGISTRO P001: ABERTURA DO BLOCO P
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_P001

    '******** registro P001 *****************
    'obrigatório se houver registros 145
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE NumIntSped = ? AND Registro = ? AND ValorRecContrPrev > 0", lContador, objEFD.lNumIntSped, "F500")
    If lErro <> AD_SQL_SUCESSO Then gError 209467

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209468
        
    objEFD.objSpedFiscal.sTipo = "P001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBlocoP = True
    Else
        iTemReg = 1
        bAbreBlocoP = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "P001")
    Call objEFD.Reg_Campo_add(2, "IND_MOV", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_P001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_P001:

   Processa_Registro_P001 = gErr

   Select Case gErr
   
        Case 209467, 209468
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209469)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_P010(ByVal objEFD As ClassEFDPisCofinsSel, dTotalContrPrev As Double) As Long
'REGISTRO P010: IDENTIFICAÇÃO DO ESTABELECIMENTO
'Este registro tem o objetivo de identificar o estabelecimento da pessoa jurídica a que se referem as operações informadas
'neste bloco. Só devem ser escriturados no Registro P010 os estabelecimentos que efetivamente tenham auferido receitas
'sujeitas à incidência da Contribuição Previdenciária sobre a Receita Bruta.
'O estabelecimento que não realizou operações passíveis de registro nesse bloco, no período da escrituração, não deve ser
'identificado no Registro P010.
'Para cada estabelecimento cadastrado em “P010”, deve ser informado nos registros de nível inferior (Registros Filho
'P100 e/ou P110) as informações necessárias para a apuração da Contribuição Previdenciária sobre Receitas
Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 2) As Long
Dim sChave As String, sCgc As String

On Error GoTo Erro_Processa_Registro_P010

    '******************* registro P010 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    sCgc = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT DISTINCT(CGC) FROM SpedDocFiscais WHERE NumIntSped = ? AND Registro = ? AND ValorRecContrPrev > 0 ORDER BY CGC", sCgc, objEFD.lNumIntSped, "F500")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209470
   
    Do While lErro = AD_SQL_SUCESSO
                
        sChave = "CNPJ\CPF: " & sCgc
        Call objEFD.objTela.SPED_Status("Processando Registro P010", sChave)
        
        objEFD.objSpedFiscal.sTipo = "P010"
        objEFD.objSpedFiscal.lSeqPai = 0
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "P010")
        Call objEFD.Reg_Campo_add(2, "CNPJ", SFF_Texto(sCgc, 14))
        Call objEFD.Finaliza_Reg
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_P100(objEFD, alComando(1), alComando(2), objEFD.objSpedFiscal.lSeq, sCgc, dTotalContrPrev)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209471
    
    Loop

    Processa_Registro_P010 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_P010:

   Processa_Registro_P010 = gErr

   Select Case gErr
   
        Case 209470, 209471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209472)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_P010", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_P100(ByVal objEFD As ClassEFDPisCofinsSel, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lSeqPai As Long, ByVal sCgc As String, dTotalContrPrev As Double) As Long
'REGISTRO P100: CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA
'Registro específico da escrituração da contribuição previdenciária incidente sobre o valor da receita bruta, prevista na
'legislação tributária, conforme a Tabela “5.1.1- Atividades, Produtos e Serviços Sujeitos à Contribuição Previdenciária
'sobre a Receita Bruta”.
'Poderão ser gerados um ou vários registros “P100” para o mesmo estabelecimento, de acordo com as chaves definidas
'para o registro. (Chaves: DT_INI + DT_FIN + COD_ATIV_ECON + ALIQ_CONT + COD_CTA).
Dim lErro As Long
Dim sChave As String, sSQL As String
Dim dTotalRec As Double, sCodAtividadeTab511 As String
Dim dAliquotaContrPrev As Double, dTotalRecAtiv As Double
Dim dValorContrPrev As Double

On Error GoTo Erro_Processa_Registro_P100

    '******** registro P100 *****************
    objEFD.objSpedFiscal.sTipo = "P100"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    lErro = Comando_Executar(lComando1, "SELECT SUM(D.ValorRecebido) FROM SpedDocFiscais AS D WHERE D.Registro = 'F500' AND D.NumIntSped = ? AND D.CGC = ?", dTotalRec, objEFD.lNumIntSped, sCgc)
    If lErro <> AD_SQL_SUCESSO Then gError 211983

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211984
    
    sCodAtividadeTab511 = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT I.CodAtividadeTab511, I.AliquotaContrPrev, SUM(I.ValorRecebido)"
    sSQL = sSQL & " " & "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL & " " & "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Registro = 'F500' AND I.NumIntSped = ? AND D.CGC = ?"
    sSQL = sSQL & " " & "GROUP BY I.CodAtividadeTab511, I.AliquotaContrPrev"
    sSQL = sSQL & " " & "ORDER BY I.CodAtividadeTab511, I.AliquotaContrPrev"
   
    lErro = Comando_Executar(lComando2, sSQL, sCodAtividadeTab511, dAliquotaContrPrev, dTotalRecAtiv, objEFD.lNumIntSped, sCgc)
    If lErro <> AD_SQL_SUCESSO Then gError 211983

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211984
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sChave = "CodAtividadeTab511: " & sCodAtividadeTab511 & " Alíquota: " & Format(dAliquotaContrPrev, "PERCENT")
    
        Call objEFD.objTela.SPED_Status("Processando Registro P100", sChave)
        
        dValorContrPrev = dTotalRecAtiv * dAliquotaContrPrev
    
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "P100")
        Call objEFD.Reg_Campo_add(2, "DT_INI", SFF_Data(objEFD.dtDataIni))
        Call objEFD.Reg_Campo_add(3, "DT_FIN", SFF_Data(objEFD.dtDataFim))
        Call objEFD.Reg_Campo_add(4, "VL_REC_TOT_EST", SFF_Num(dTotalRec))
        Call objEFD.Reg_Campo_add(5, "COD_ATIV_ECON", SFF_Texto(sCodAtividadeTab511))
        Call objEFD.Reg_Campo_add(6, "VL_REC_ATIV_ESTAB", SFF_Num(dTotalRecAtiv))
        Call objEFD.Reg_Campo_add(7, "VL_EXC", 0)
        Call objEFD.Reg_Campo_add(8, "VL_BC_CONT", SFF_Num(dTotalRecAtiv))
        Call objEFD.Reg_Campo_add(9, "ALIQ_CONT", SFF_Perc(dAliquotaContrPrev))
        Call objEFD.Reg_Campo_add(10, "VL_CONT_APU", SFF_Num(dValorContrPrev))
        Call objEFD.Reg_Campo_add(11, "COD_CTA", "")
        Call objEFD.Reg_Campo_add(12, "INFO_COMPL", "")
        Call objEFD.Finaliza_Reg
        
        dTotalContrPrev = dTotalContrPrev + dValorContrPrev
                        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211984

    Loop
    
    Processa_Registro_P100 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_P100:

   Processa_Registro_P100 = gErr

   Select Case gErr
   
        Case 211983, 211984
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211985)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_P100", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_P200(ByVal objEFD As ClassEFDPisCofinsSel, dTotalContrPrev As Double) As Long
'REGISTRO P200: CONSOLIDAÇÃO DA CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A
'RECEITA BRUTA
'Registro de consolidação da contribuição previdenciária incidente sobre o valor da receita bruta, devida pela empresa no
'período, correspondente ao somatório da contribuição sobre a receita bruta mensal de cada estabelecimento, apurada no
'Registro “P100”.
Dim lErro As Long
Dim sChave As String, sSQL As String
Dim dTotalRec As Double, sCodAtividadeTab511 As String
Dim dAliquotaContrPrev As Double, dTotalRecAtiv As Double
Dim dValorContrPrev As Double

On Error GoTo Erro_Processa_Registro_P200

    '******** registro P200 *****************
    objEFD.objSpedFiscal.sTipo = "P200"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    Call objEFD.objTela.SPED_Status("Processando Registro P200", sChave)
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "P200")
    Call objEFD.Reg_Campo_add(2, "PER_REF", Format(objEFD.dtDataIni, "MMYYYY"))
    Call objEFD.Reg_Campo_add(3, "VL_TOT_CONT_APU", SFF_Num(dTotalContrPrev))
    Call objEFD.Reg_Campo_add(4, "VL_TOT_AJ_REDUC", 0)
    Call objEFD.Reg_Campo_add(5, "VL_TOT_AJ_ACRES", 0)
    Call objEFD.Reg_Campo_add(6, "VL_TOT_CONT_DEV", SFF_Num(dTotalContrPrev))
    
    If objEFD.objFilialEmpresa.iTipoTribAtividade = 2 Then
        Call objEFD.Reg_Campo_add(7, "COD_REC", "299101") '2991 – Contribuição Previdenciária Sobre Receita Bruta – Indústria
    Else
        Call objEFD.Reg_Campo_add(7, "COD_REC", "298501") '2985 – Contribuição Previdenciária Sobre Receita Bruta – Serviços
    End If
    Call objEFD.Finaliza_Reg
        
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
    Processa_Registro_P200 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_P200:

   Processa_Registro_P200 = gErr

   Select Case gErr
   
        Case 211983, 211984
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211985)
    
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_P200", sChave)

    Exit Function
    
End Function

Private Function Processa_Registro_P990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO A990: ENCERRAMENTO DO BLOCO P
'Este registro destina-se a identificar o encerramento do bloco A e informar a quantidade de linhas (registros)
'existentes no bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_P990

    '******************* registro P990 **********************************
    objEFD.objSpedFiscal.sTipo = "P990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1 ' ele mesmo
        
    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "P" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "P990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_P", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_P990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_P990:

   Processa_Registro_P990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209478)
    
    End Select
    
    Exit Function
    
End Function

