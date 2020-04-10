VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABPagRem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function BorderosPagto_Criar_ArquivoCNAB(ByVal objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
   
Dim objCCI As New ClassContasCorrentesInternas
Dim lErro As Long

On Error GoTo Erro_BorderosPagto_Criar_ArquivoCNAB

    'Lê a conta corrente do bordero de pagamento
    lErro = CF("ContaCorrenteInt_Le", objGeracaoArqCNABPag.objBorderoPagto.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51750
    If lErro <> SUCESSO Then Error 51751
   
    Set objGeracaoArqCNABPag.objCCI = objCCI
   
   'Seleciona o banco
    Select Case objCCI.iCodBanco
    
'        Case 1
'            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_BcoBrasil", objGeracaoArqCNABPag)
'            If lErro <> SUCESSO Then Error 62424
    
        Case 237
            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_Bradesco", objGeracaoArqCNABPag)
            If lErro <> SUCESSO Then Error 62425

        Case 275
            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_BcoReal", objGeracaoArqCNABPag)
            If lErro <> SUCESSO Then Error 62424
    
'        Case 341
'            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_Itau", objGeracaoArqCNABPag)
'            If lErro <> SUCESSO Then Error 51752
 
        Case 409
            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_Unibanco", objGeracaoArqCNABPag)
            If lErro <> SUCESSO Then Error 62426
        
        'cnab 240
        '356: real abn amro
        Case 1, 33, 341, 356, 399, 353, 8
            lErro = CF("BorderoPagto_Criar_ArquivoCNAB_240", objGeracaoArqCNABPag)
            If lErro <> SUCESSO Then gError 62071
            
        Case Else
            Error 59356
                
    End Select

    BorderosPagto_Criar_ArquivoCNAB = SUCESSO
        
    Exit Function

Erro_BorderosPagto_Criar_ArquivoCNAB:

    BorderosPagto_Criar_ArquivoCNAB = Err
    
    Select Case Err
    
        Case 51750, 51752, 62424, 62425, 62426
        
        Case 51751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objCCI.iCodigo)
        
        Case 59356
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDPAGTO_BANCO_NAO_TRATADO", Err, objCCI.iCodBanco)
            '??? falta colocar msg de erro
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147307)
            
    End Select
    
    Exit Function

End Function

Function PagamentoEletronico_Gera_NomeArquivo(sNomeArquivo As String, ByVal objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Gera o nome do arquivo de acordo com a data e o sequencial do bordero

Dim lErro As Long
Dim lSequencialArq As Long
Dim sNomeAux As String, sValorInformacao As String
Dim bGerou As Boolean, lComando1 As Long, lComando2 As Long, lComando3 As Long

On Error GoTo Erro_PagamentoEletronico_Gera_NomeArquivo

    bGerou = False
    lSequencialArq = 0
    
    Do While Not bGerou
    
        lSequencialArq = lSequencialArq + 1
        'Gera o nome do arquivo\
        sNomeAux = "PG" & Format(gdtDataHoje, "DDMM") & FormataCpoNum(lSequencialArq, 2) & ".REM"
        
        If Len(Trim(Dir(objGeracaoArqCNABPag.sDiretorio & "\" & sNomeAux))) = 0 Then bGerou = True
    
    Loop
        
    sNomeArquivo = objGeracaoArqCNABPag.sDiretorio & "\" & sNomeAux

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 130418
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 130419
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 130419
    
    sValorInformacao = String(255, 0)
    lErro = Comando_ExecutarPos(lComando1, "SELECT ValorInformacao FROM CNABPagConfig WHERE FilialEmpresa = ? AND CodContaCorrenteInterna = ? AND CodBanco = ? AND InfoCodigo = ?", 0, _
        sValorInformacao, EMPRESA_TODA, 0, objGeracaoArqCNABPag.objCCI.iCodBanco, 1)
    If lErro <> AD_SQL_SUCESSO Then gError 130420
    
    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130421
    If lErro <> AD_SQL_SUCESSO Then 'gError 130424
    
        lErro = Comando_Executar(lComando2, "INSERT INTO CNABPagConfig (FilialEmpresa,CodContaCorrenteInterna,CodBanco,InfoCodigo,InfoTexto,InfoDescricao,ValorInformacao) VALUES (?,?,?,?,?,?,?) ", _
        0, 0, objGeracaoArqCNABPag.objCCI.iCodBanco, 1, "Prox. Seq. do Arquivo", "Próximo Sequencial do Arquivo de Remessa", "2")
        If lErro <> AD_SQL_SUCESSO Then gError 130423
    
        lSequencialArq = 1
    
    Else
    
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 130422
        
        lSequencialArq = StrParaLong(sValorInformacao)
        sValorInformacao = CStr(lSequencialArq + 1)
        
        lErro = Comando_ExecutarPos(lComando2, "UPDATE CNABPagConfig SET ValorInformacao = ?", lComando1, sValorInformacao)
        If lErro <> AD_SQL_SUCESSO Then gError 130423
    
    End If
    
    objGeracaoArqCNABPag.lSequencialArq = lSequencialArq
    
    'obtem o codigo da empresa no banco
    sValorInformacao = String(255, 0)
    lErro = Comando_ExecutarPos(lComando3, "SELECT ValorInformacao FROM CNABPagConfig WHERE FilialEmpresa = ? AND CodContaCorrenteInterna = ? AND CodBanco = ? AND InfoCodigo = ?", 0, _
        sValorInformacao, EMPRESA_TODA, 0, objGeracaoArqCNABPag.objCCI.iCodBanco, 2)
    If lErro <> AD_SQL_SUCESSO Then gError 130531
    
    lErro = Comando_BuscarProximo(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130532
    If lErro = AD_SQL_SUCESSO Then objGeracaoArqCNABPag.sCodEmpresaBanco = sValorInformacao
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    PagamentoEletronico_Gera_NomeArquivo = SUCESSO

    Exit Function

Erro_PagamentoEletronico_Gera_NomeArquivo:

    PagamentoEletronico_Gera_NomeArquivo = gErr
    
    Select Case gErr
    
        Case 130318, 130419
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130420, 130421, 130424
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CNABPAGCONFIG", gErr)
        
        Case 130422
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_CNABPAGCONFIG", gErr)
        
        Case 130423
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CNABPAGCONFIG", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147308)
            
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function
    
End Function

Public Function BorderoPagto_Abre_TelaRemessaArq(ByVal sNomeArqParam As String, ByVal sDiretorio As String, ByVal objBorderoPagtoEmiss As ClassBorderoPagEmissao) As Long
    
Dim lErro As Long

On Error GoTo Erro_BorderoPagto_Abre_TelaRemessaArq

    BorderoPagGeraArqRem.sNomeArqParam = sNomeArqParam
    
    Call BorderoPagGeraArqRem.Trata_Parametros(sDiretorio, objBorderoPagtoEmiss)
    
    BorderoPagGeraArqRem.Show vbModal

    Set BorderoPagGeraArqRem = Nothing

    BorderoPagto_Abre_TelaRemessaArq = SUCESSO
    
    Exit Function

Erro_BorderoPagto_Abre_TelaRemessaArq:

    BorderoPagto_Abre_TelaRemessaArq = Err
    
    Select Case Err
    
        Case 62422, 62423
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147309)
            
    End Select
    
''    Call Rotina_ErrosBatch2("Geração de Arquivo de Pagamentos")
    
    Exit Function
    
End Function
