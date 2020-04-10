VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassExtrCNAB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const EXTR_TAM_REG = 200

Private Const CNAB_SUCESSO = 0
Private Const CNAB_BCO_INVALIDO = 1
Private Const CNAB_DRIVER = 2
Private Const CNAB_ABRIR_ARQ = 3
Private Const CNAB_MEMORIA = 4
Private Const CNAB_CONFIG_DRIVER = 5
Private Const CNAB_ERRO_PARAM = 6
Private Const CNAB_ERRO_INTERNO = 7
Private Const CNAB_ERRO_LER_ARQ = 8
Private Const CNAB_ERRO_SEQUENCIA = 9
Private Const CNAB_ERRO_BCO = 10
Private Const CNAB_ERRO_SERVICO = 11

Private Const CNAB_EXT_NUM_CTAS = 12       '# de ctas no trailer <> do num de ctas encontrado
Private Const CNAB_EXT_TOT_DEB = 13         '<> entre o total do trailer e dos lctos
Private Const CNAB_EXT_TOT_CRE = 14        '<> entre o total do trailer e dos lctos

Private Const CNABEXT_LER_OK = 0
Private Const CNABEXT_LER_ACABOU_ARQ = 1
Private Const CNABEXT_LER_ERRO = 2

' tipos de registro fisico
Private Const CNABEXT_REG_HEADER = 0
Private Const CNABEXT_REG_SALDOINICIAL = 1
Private Const CNABEXT_REG_LCTO = 2
Private Const CNABEXT_REG_SALDOFINAL = 3
Private Const CNABEXT_REG_TRAILER = 4
Private Const CNABEXT_REG_INDIC_LCTO_FUTURO = 5
Private Const CNABEXT_REG_LCTO_FUTURO = 6
Private Const CNABEXT_REG_LCTO_SIMULADO = 7


'??? retirar funcoes de I/O daqui p/rotinas alguma coisa ???
'??? talvez toda a parte CNAB possa estar em um vbp separado "fechado"

Private Declare Function Extrato_AbrirInt Lib "ADCNAB.DLL" Alias "Extrato_Abrir" (ByVal iBanco As Integer, ByVal sNomeArq As String, l_num_regs As Long, lExtrato As Long) As Long
'cria estrutura p/recepcao do extrato
'E: iBanco : Ex; Bradesco 237, Unibanco 409, BB: 001,...
'E: sNomeArq: nome do arquivo contendo o extrato de uma ou mais contas
'S: lExtrato: ponteiro p/objeto que vai "gerenciar" o processo.
'Retorno: 0 sucesso, 1: bco invalido, 2, nao conseguiu abrir o arquivo, 3: lpfnCallback invalida, 4:lplpExtrato invalido

Private Declare Function Extrato_Ler_RegistroInt Lib "ADCNAB.DLL" Alias "Extrato_Ler_Registro" (ByVal lExtrato As Long, iTipo_reg As Integer, ByVal sReg As String) As Long
'ler um registro do arquivo
Private Declare Function Extrato_FecharInt Lib "ADCNAB.DLL" Alias "Extrato_Fechar" (ByVal lExtrato As Long) As Long
'libera recursos alocados ao termino da operacao


Private lExtrato As Long            '' ponteiro p/contexto mantido por adcnab.dll
Private mvariBanco As Integer           '' codigo do bco. Ex. Bradesco : 237
Private lNumRegsArq As Long         '' # de registros total do arquivo
Private lNumRegsLidos As Long       '' # de registros lidos

Private iPulaConta As Integer '''pula cta que nao esta cadastrada no sistema

Private objExtrBcoCab As New ClassExtrBcoCab
Private objExtrBcoDet As New ClassExtrBcoDet

Private lTransacao As Long
Private lComando1 As Long, lComando2 As Long '' para uso na insercao de lctos
Private lComando3 As Long, lComando4 As Long '' para insercao de extratos

Public Function Extrato_Abrir(iBanco As Integer, sNomeArq As String) As Long
'inicializacao da leitura de um arquivo com extrato bancario p/conciliacao
Dim lErro As Long

On Error GoTo Erro_Extrato_Abrir

    mvariBanco = iBanco

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 7141

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 7142

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 7143

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 7143

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 7144

    lErro = Extrato_AbrirInt(iBanco, sNomeArq, lNumRegsArq, lExtrato)
    If lErro <> SUCESSO Then Error 7145

    objExtrBcoCab.sNomeArquivo = sNomeArq
    objExtrBcoCab.iBanco = iBanco
    objExtrBcoDet.iConciliado = 0

    Extrato_Abrir = SUCESSO

    Exit Function

Erro_Extrato_Abrir:

    Extrato_Abrir = Err

    Select Case Err
        Case 7141, 7142, 7143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 7144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 7145
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149721)

    End Select

    Exit Function

End Function

Private Function Extrato_TrataHeader(sReg As String) As Long
 ''      1 a 6: data de gravacao

    objExtrBcoCab.dtDataGravacao = Extr_ConvData(Mid(sReg, 1, 6))

    Extrato_TrataHeader = SUCESSO

End Function

Private Function Extrato_TrataTrailer(sReg As String) As Long
Dim lErro As Long

On Error GoTo Erro_Extrato_TrataTrailer
    'só vai poder fazer o commit depois de ler o registro trailer pois
    'nele estarao os dados p/verificacao da integridade do arquivo.

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 7146

    lTransacao = 0

    Extrato_TrataTrailer = SUCESSO

    Exit Function

Erro_Extrato_TrataTrailer:

    Extrato_TrataTrailer = Err

    Select Case Err

        Case 7146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149722)

    End Select

    Exit Function

End Function

Private Function Extrato_TrataLcto(sReg As String) As Long

Dim lErro As Long

On Error GoTo Erro_Extrato_TrataLcto

''   1 a 3: ach_categoria[ 3 ];             // 9(3)
''   4 a 7: ach_cod_historico[ 4 ];         // 9(4)
''  8 a 32: ach_descr_historico[ 25 ];      //  X(25)
''  33 a 38: ach_num_doc[ 6 ];              //  9(6)
''  39 a 44: ach_data[ 6 ];                 // 9(6) ddmmaa
''  45 a 62:  char ach_valor[ 18 ];         // 9(16)V99
''  63 a 63:  char ch_DebCred;              // "D" ou "C"
''  64 a 64:  char ch_CPMF;                 // "S" ou "N"
''  65 a 70: seq reg                        // 9(6)

    objExtrBcoDet.lSeqLcto = CLng(Mid(sReg, 65, 6))
    objExtrBcoDet.iCategoria = CInt(Mid(sReg, 1, 3))
    objExtrBcoDet.sCodLctoBco = Mid(sReg, 4, 4)
    objExtrBcoDet.sHistorico = Mid(sReg, 8, 25)
    objExtrBcoDet.sDocumento = Mid(sReg, 33, 6)
    objExtrBcoDet.dtData = Extr_ConvData(Mid(sReg, 39, 6))
    objExtrBcoDet.dValor = Extr_ConvValMoeda(Mid(sReg, 45, 18), Mid(sReg, 63, 1), 2)
    objExtrBcoDet.sIncideCPMF = Mid(sReg, 64, 1)

    'incluir o lcto se este ainda nao tiver sido registrado
    lErro = CF("ExtrBcoDet_Inclui", objExtrBcoDet, lComando1, lComando2)
    If lErro <> SUCESSO Then Error 7147

    Extrato_TrataLcto = SUCESSO

    Exit Function

Erro_Extrato_TrataLcto:

    Extrato_TrataLcto = Err

    Select Case Err

        Case 7147
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149723)

    End Select

    Exit Function

End Function

Private Function Extrato_TrataSaldoFinal(sReg As String) As Long

Dim lErro As Long
''  1 a 6: ach_data[ 6 ];             // 9(6)     ddmmaa
'' 7 a 24: ach_valor[ 18 ];           // 9(16)V99
'' 25 a 25: ch_DebCred;                // ("D" ou "C")

On Error GoTo Erro_Extrato_TrataSaldoFinal

    objExtrBcoCab.dtDataSaldoFinal = Extr_ConvData(Mid(sReg, 1, 6))
    objExtrBcoCab.dSaldoFinal = Extr_ConvValMoeda(Mid(sReg, 7, 18), Mid(sReg, 25, 1), 2)

    'incluir registro do extrato
    lErro = CF("ExtrBcoCab_Grava_Batch", objExtrBcoCab, lComando3, lComando4)
    If lErro <> SUCESSO Then Error 7148

    Extrato_TrataSaldoFinal = SUCESSO

    Exit Function

Erro_Extrato_TrataSaldoFinal:

    Extrato_TrataSaldoFinal = Err

    Select Case Err
        Case 7148
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149724)

    End Select

    Exit Function

End Function

Private Function Extrato_TrataSaldoInicial(sReg As String) As Long
''     1 a 4: Agencia[ 4 ];          // 9(4)
''     5 a 12 ach_Conta[ 8 ];            // 9(8)     (ver no bbd como obter dentro de um cpo de 20 bytes)
''    13 a 17 char ach_num_arquivo[ 5 ];      // 9(5)     (sequencial p/conta)
''    18 a 23 char ach_data[ 6 ];             // 9(6)     ddmmaa
''    24 a 41 char ach_valor[ 18 ];           // 9(16)V99
''    42 a 42 char ch_DebCred;                // ("D" ou "C")
Dim sAgencia As String
Dim sConta As String
Dim iCodConta As Integer
Dim lErro As Long

On Error GoTo Erro_Extrato_TrataSaldoInicial

    'obtem conta corrente interna
    sAgencia = Mid(sReg, 1, 4)
    sConta = Mid(sReg, 5, 8)
    lErro = CF("CtaCorrente_Procura", mvariBanco, sAgencia, sConta, iCodConta)
    If lErro <> SUCESSO Then Error 7166
    
    'se nao achar...
    If iCodConta = 0 Then
        ' pular a conta inteira
        iPulaConta = 1
    Else
        'armazena dados do extrato
            objExtrBcoCab.iCodConta = iCodConta
            objExtrBcoCab.iNumExtrato = CInt(Mid(sReg, 13, 5))
            objExtrBcoCab.dtDataSaldoInicial = Extr_ConvData(Mid(sReg, 18, 6))
            objExtrBcoCab.dSaldoInicial = Extr_ConvValMoeda(Mid(sReg, 24, 18), Mid(sReg, 42, 1), 2)
            objExtrBcoCab.dHoraImport = CDbl(Time)
            objExtrBcoCab.dtDataImport = Date

        'armazena dados gerais p/os lctos desta conta
            objExtrBcoDet.iCodConta = iCodConta
            objExtrBcoDet.iNumExtrato = CInt(Mid(sReg, 13, 5))

        iPulaConta = 0
    End If

    Extrato_TrataSaldoInicial = SUCESSO

    Exit Function

Erro_Extrato_TrataSaldoInicial:

    Extrato_TrataSaldoInicial = Err

    Select Case Err
    
        Case 7166
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149725)

    End Select

    Exit Function

End Function

Public Function Extrato_LerReg(iPerc As Integer, iAcabou As Integer) As Long
Dim lErro As Long, iTipo_reg As Integer, sReg As String

On Error GoTo Erro_Extrato_LerReg

    iAcabou = 0
    sReg = String(EXTR_TAM_REG, 0)

    lErro = Extrato_Ler_RegistroInt(lExtrato, iTipo_reg, sReg)
    If lErro <> SUCESSO Then Error 7149

    lNumRegsLidos = lNumRegsLidos + 1

    Select Case iTipo_reg

        Case CNABEXT_REG_HEADER:
            lErro = Extrato_TrataHeader(sReg)
            If lErro <> SUCESSO Then Error 7150

        Case CNABEXT_REG_SALDOINICIAL
            lErro = Extrato_TrataSaldoInicial(sReg)
            If lErro <> SUCESSO Then Error 7151

        Case CNABEXT_REG_LCTO:
            If iPulaConta <> 1 Then
                lErro = Extrato_TrataLcto(sReg)
                If lErro <> SUCESSO Then Error 7152
            End If

        Case CNABEXT_REG_SALDOFINAL:
            If iPulaConta <> 1 Then
                lErro = Extrato_TrataSaldoFinal(sReg)
                If lErro <> SUCESSO Then Error 7153
            End If

        Case CNABEXT_REG_TRAILER:
            lErro = Extrato_TrataTrailer(sReg)
            If lErro <> SUCESSO Then Error 7154

            iAcabou = 1

    End Select

    iPerc = CInt(100 * lNumRegsLidos / lNumRegsArq)

    Extrato_LerReg = SUCESSO

    Exit Function

Erro_Extrato_LerReg:

    Extrato_LerReg = Err

    Select Case Err
        Case 7149, 7150, 7151, 7152, 7153, 7154
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149726)

    End Select

    Exit Function

End Function

Public Function Extrato_Fechar() As Long
'termino da leitura de um arquivo com extrato bancario p/conciliacao
Dim lErro As Long

On Error GoTo Erro_Extrato_Fechar

    If lTransacao <> 0 Then Call Transacao_Rollback

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    If lExtrato <> 0 Then
    
        lErro = Extrato_FecharInt(lExtrato)
        If lErro <> SUCESSO Then Error 7155
    
        lExtrato = 0

    End If
    
    Extrato_Fechar = SUCESSO

    Exit Function

Erro_Extrato_Fechar:

    Extrato_Fechar = Err

    Select Case Err
        Case 7155
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149727)

    End Select

    If lTransacao <> 0 Then Call Transacao_Rollback

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Exit Function

End Function

Private Sub Class_Initialize()
    lExtrato = 0
    lTransacao = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0
    lComando4 = 0
End Sub

Private Sub Class_Terminate()
    
    Call Extrato_Fechar
    
End Sub

Private Function Extr_ConvData(sddmmaa As String) As Date
'retorna uma data correspondente ao string no formato ddmmaa

    Extr_ConvData = CDate(Mid(sddmmaa, 1, 2) & "/" & Mid(sddmmaa, 3, 2) & "/" & Mid(sddmmaa, 5, 2))
    
End Function

Private Function Extr_ConvValMoeda(sVal As String, sSinal As String, iNumDecImplicitas As Integer) As Double

Dim dValorConvertido As Double

    dValorConvertido = CDbl(sVal)
    If sSinal = "D" Then dValorConvertido = -dValorConvertido
    If iNumDecImplicitas > 1 Then dValorConvertido = dValorConvertido / (10 ^ iNumDecImplicitas)
    
    Extr_ConvValMoeda = dValorConvertido
    
End Function
