VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFiscalItens_Le_Customizado(ByVal objNFiscal As ClassNFiscal) As Long
    NFiscalItens_Le_Customizado = CF("NFiscal_Le_Customizado", objNFiscal)
End Function

Function NFiscal_Le_Customizado(ByVal objNFiscal As ClassNFiscal) As Long
'Lê os valores customizados dos ítens da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim lComando As Long
Dim objNFiscalMgz As New ClassNFiscalMgz
Dim sProcesso As String
Dim sID As String
Dim sContrato As String

On Error GoTo Erro_NFiscal_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141975
    
    sProcesso = String(255, 0)
    sID = String(255, 0)
    sContrato = String(255, 0)

    'Pesquisa os ítens vinculados à Nota Fiscal
    lErro = Comando_Executar(lComando, "SELECT Contrato, IDCorrespondencia, Processo  FROM NFiscal WHERE NumIntDoc = ? ", sContrato, sID, sProcesso, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 141976

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141977

    If lErro = AD_SQL_SUCESSO Then
        
        objNFiscalMgz.sContrato = sContrato
        objNFiscalMgz.sIDCorrespondencia = sID
        objNFiscalMgz.sProcesso = sProcesso
        
        Set objNFiscal.objInfoUsu = objNFiscalMgz

    End If

    NFiscal_Le_Customizado = SUCESSO

    Exit Function

Erro_NFiscal_Le_Customizado:

    NFiscal_Le_Customizado = gErr

    Select Case gErr

        Case 141975
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141976, 141977
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141978)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ContratoMgz_Le(ByVal objContrato As ClassContratoMgz) As Long
'preenche objContrato lendo dados as tabelas Contratos e ProcessosContrato
'objContrato.sContrato e objContrato.lCliente tem que estar preenchidos

Dim lErro As Long, alComando(1 To 2) As Long, iIndice As Integer, iVendedor As Integer
Dim tProc As typeProcessoContrato, objProc As ClassProcContratoMgz

On Error GoTo Erro_ContratoMgz_Le

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 99999
    Next

    lErro = Comando_Executar(alComando(1), "SELECT Vendedor FROM ContratosMgz WHERE Contrato = ? AND Cliente = ?", iVendedor, objContrato.sContrato, objContrato.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 99999

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99999
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_OBJETO_NAO_CADASTRADO
    
    objContrato.iVendedor = iVendedor
    
    Set objContrato.colProcessos = New Collection

    With tProc
        
        .sProcesso = String(STRING_PROCESSO_CONTRATO_ID, 0)
        .sDescricao = String(STRING_PROCESSO_CONTRATO_DESCRICAO, 0)
        .sObservacao = String(STRING_PROCESSO_CONTRATO_OBSERVACAO, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT Seq, Processo, Tipo, Descricao, Valor, DataCobranca, Observacao FROM ProcessosContrato WHERE Contrato = ? AND Cliente = ? ORDER BY Seq", _
            .iSeq, .sProcesso, .iTipo, .sDescricao, .dValor, .dtDataCobranca, .sObservacao, objContrato.sContrato, objContrato.lCliente)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_OBJETO_NAO_CADASTRADO

    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99999
        
    Do While lErro = AD_SQL_SUCESSO
            
        Set objProc = New ClassProcContratoMgz
    
        With objProc
    
            .dtDataCobranca = tProc.dtDataCobranca
            .dValor = tProc.dValor
            .iSeq = tProc.iSeq
            .iTipo = tProc.iTipo
            .lCliente = objProc.lCliente
            .sContrato = objProc.sContrato
            .sDescricao = tProc.sDescricao
            .sObservacao = tProc.sObservacao
            .sProcesso = tProc.sProcesso
        
        End With
        
        objContrato.colProcessos.Add objProc
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99999
    
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ContratoMgz_Le = SUCESSO
     
    Exit Function
    
Erro_ContratoMgz_Le:

    ContratoMgz_Le = gErr
     
    Select Case gErr
          
        Case 99999
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROCESSOCONTRATO", gErr)
          
        Case ERRO_OBJETO_NAO_CADASTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ProcessoContrato_Le(ByVal objProc As ClassProcContratoMgz) As Long
'retorna os dados do processo.
'ATENCAO: Se objProc.iSeq for zero, retorna os dados do menor seq referente ao processo

Dim lErro As Long, lComando As Long, tProc As typeProcessoContrato

On Error GoTo Erro_ProcessoContrato_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99999
    
    With tProc
        
        .sDescricao = String(STRING_PROCESSO_CONTRATO_DESCRICAO, 0)
        .sObservacao = String(STRING_PROCESSO_CONTRATO_OBSERVACAO, 0)
        
        If objProc.iSeq <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT Seq, Tipo, Descricao, Valor, DataCobranca, Observacao FROM ProcessosContrato WHERE Contrato = ? AND Cliente = ? AND Processo = ? AND Seq = ?", _
                .iSeq, .iTipo, .sDescricao, .dValor, .dtDataCobranca, .sObservacao, objProc.sContrato, objProc.lCliente, objProc.sProcesso, objProc.iSeq)
        Else
            lErro = Comando_Executar(lComando, "SELECT Seq, Tipo, Descricao, Valor, DataCobranca, Observacao FROM ProcessosContrato WHERE Contrato = ? AND Cliente = ? AND Processo = ? ORDER BY Seq", _
                .iSeq, .iTipo, .sDescricao, .dValor, .dtDataCobranca, .sObservacao, objProc.sContrato, objProc.lCliente, objProc.sProcesso)
        End If
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 99999

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99999
        
    If lErro <> SUCESSO Then gError ERRO_OBJETO_NAO_CADASTRADO
    
    With objProc
    
        .dtDataCobranca = tProc.dtDataCobranca
        .dValor = tProc.dValor
        .iSeq = tProc.iSeq
        .iTipo = tProc.iTipo
        .sDescricao = tProc.sDescricao
        .sObservacao = tProc.sObservacao
    
    End With
    
    Call Comando_Fechar(lComando)
    
    ProcessoContrato_Le = SUCESSO
     
    Exit Function
    
Erro_ProcessoContrato_Le:

    ProcessoContrato_Le = gErr
     
    Select Case gErr
    
        Case 99999
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROCESSOCONTRATO", gErr)
          
        Case ERRO_OBJETO_NAO_CADASTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

