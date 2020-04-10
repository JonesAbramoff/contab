VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSelectTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TRPRateio_Le(ByVal colTRPRateio As Collection) As Long
'le os lançamentos relativos ao TRPRateio

Dim lComando As Long
Dim lErro As Long
Dim objTRPRateio As ClassTRPRateio
Dim sConta As String
Dim sCcl As String
Dim sDescricao As String
Dim iPeriodo As Integer
Dim iExercicio As Integer


On Error GoTo Erro_TRPRateio_Le
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197622

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)
    sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Exercicio, Periodo, Conta, Ccl, Descricao  FROM TRPRateio ORDER BY Seq", iExercicio, iPeriodo, sConta, sCcl, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 197623

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197624

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTRPRateio = New ClassTRPRateio

        objTRPRateio.iExercicio = iExercicio
        objTRPRateio.iPeriodo = iPeriodo
        objTRPRateio.sConta = sConta
        objTRPRateio.sCcl = sCcl
        objTRPRateio.sDescricao = sDescricao

        colTRPRateio.Add objTRPRateio

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197625
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    TRPRateio_Le = SUCESSO
    
    Exit Function
 
Erro_TRPRateio_Le:

    TRPRateio_Le = gErr

    Select Case gErr
    
        Case 197622
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 197623 To 197625
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPRATEIO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197626)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

