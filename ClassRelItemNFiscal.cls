VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelItemNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltNaturezaNF As String
Private dtUltDataNF As Date
Private lUltNumeroNF As Long
Private sUltSerieNF As String
Private lUltNumIntDoc As Long

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

'Maristela(inicio)
Public Function ObterNaturezaNF(lNumIntDoc As Long, sNaturezaNF As String) As Long
'Para obter a natureza da Nota Fiscal

Dim lErro As Long

On Error GoTo Erro_ObterNaturezaNF
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelNF_NatDataNumSerie(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90450
    
    End If
    
    sNaturezaNF = sUltNaturezaNF
    
    ObterNaturezaNF = SUCESSO
    
    Exit Function
    
Erro_ObterNaturezaNF:

    ObterNaturezaNF = gErr
    
    Select Case gErr

        Case 90450
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153180)

    End Select

    Exit Function

End Function

Public Function ObterDataNF(lNumIntDoc As Long, dtDataNF As Date) As Long
'Para obter a Data da Nota Fiscal

Dim lErro As Long

On Error GoTo Erro_ObterDataNF
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelNF_NatDataNumSerie(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90451
    
    End If
    
    dtDataNF = dtUltDataNF
    
    ObterDataNF = SUCESSO
    
    Exit Function
    
Erro_ObterDataNF:

    ObterDataNF = gErr
    
    Select Case gErr

        Case 90451
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153181)

    End Select

    Exit Function

End Function

Public Function ObterNumeroNF(lNumIntDoc As Long, lNumeroNF As Long) As Long
'Para obter o Numero da Nota Fiscal

Dim lErro As Long

On Error GoTo Erro_ObterNumeroNF
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelNF_NatDataNumSerie(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90452
    
    End If
    
    lNumeroNF = lUltNumeroNF
    
    ObterNumeroNF = SUCESSO
    
    Exit Function
    
Erro_ObterNumeroNF:

    ObterNumeroNF = gErr
    
    Select Case gErr

        Case 90452
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153182)

    End Select

    Exit Function

End Function

Public Function ObterSerieNF(lNumIntDoc As Long, sSerieNF As String) As Long
'Para obter a Serie da Nota Fiscal

Dim lErro As Long

On Error GoTo Erro_ObterSerieNF
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelNF_NatDataNumSerie(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90453
    
    End If
    
    sSerieNF = sUltSerieNF
    
    ObterSerieNF = SUCESSO
    
    Exit Function
    
Erro_ObterSerieNF:

    ObterSerieNF = gErr
    
    Select Case gErr

        Case 90453
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153183)

    End Select

    Exit Function

End Function

Private Function ObterRelNF_NatDataNumSerie(lNumIntDoc As Long) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim sNaturezaNF As String
Dim dtDataNF As Date
Dim lNumeroNF As Long
Dim sSerieNF As String

On Error GoTo Erro_ObterRelNF_NatDataNumSerie

    lErro = Obter_NatDataNumSerie_NumIntDoc(lNumIntDoc, sNaturezaNF, dtDataNF, lNumeroNF, sSerieNF, lComando)
    If lErro <> SUCESSO Then gError 90454
    
    lUltNumIntDoc = lNumIntDoc
    
    sUltNaturezaNF = sNaturezaNF
    dtUltDataNF = dtDataNF
    lUltNumeroNF = lNumeroNF
    sUltSerieNF = sSerieNF
    
    ObterRelNF_NatDataNumSerie = SUCESSO
    
    Exit Function
    
Erro_ObterRelNF_NatDataNumSerie:

    ObterRelNF_NatDataNumSerie = gErr
    
    Select Case gErr

        Case 90454
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153184)

    End Select

    Exit Function
    
End Function

Function Obter_NatDataNumSerie_NumIntDoc(lNumIntDoc As Long, sNaturezaNF As String, dtDataNF As Date, lNumeroNF As Long, sSerieNF As String, lComando As Long) As Long
'Obtem a Natureza, Data, Numero, serie da Nota fiscal a partir do NumIntDoc da tabela ItensNFiscal(passada por parametro)

Dim lErro As Long

On Error GoTo Erro_Obter_NatDataNumSerie_NumIntDoc

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90455
    End If
    
    sNaturezaNF = String(STRING_NATUREZAOP_CODIGO, 0)
    sSerieNF = String(STRING_SERIE, 0)
    
    'Lê da tabela ItensNFiscal o NumIntDoc passado por parametro e faz o relacionamento com a tabela NFiscal
    lErro = Comando_Executar(lComando, "SELECT TributacaoItemNF.NaturezaOp, NFiscal.DataEmissao, NFiscal.NumNotaFiscal, NFiscal.Serie " & _
            "FROM TributacaoItemNF, ItensNFiscal, NFiscal WHERE TributacaoItemNF.NumIntDoc = ItensNFiscal.NumIntDoc AND ItensNFiscal.NumIntDoc = ? AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc ", sNaturezaNF, dtDataNF, lNumeroNF, sSerieNF, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 90456
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90457

    Obter_NatDataNumSerie_NumIntDoc = SUCESSO

    Exit Function

Erro_Obter_NatDataNumSerie_NumIntDoc:

    Obter_NatDataNumSerie_NumIntDoc = gErr

    Select Case gErr

        Case 90455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90456, 90457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153185)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Maristela(fim)

