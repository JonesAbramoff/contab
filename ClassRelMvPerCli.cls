VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvPerCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim lComando As Long
Dim lComando2 As Long
Dim iFilialEmpresaAnt As Integer
Dim lClienteAnt As Long
Dim dtDataAnt As Date
Dim dSaldoAnt As Double

Function Cliente_Obtem_SaldoAnterior(dSaldo As Double, lCliente As Long, dtData As Date, iFilialEmpresa As Integer) As Long
'Retorna em dSaldo o SaldoContabil a receber do Cliente, consultando as tabelas MvPerCli e MvDiaCli

Dim lErro As Long
Dim iMes As Integer
Dim iAno As Integer
Dim dSaldoDiario As Double

On Error GoTo Erro_Cliente_Obtem_SaldoAnterior

    If iFilialEmpresa <> iFilialEmpresaAnt Or dtData <> dtDataAnt Or lCliente <> lClienteAnt Then
            
        'Busca dSaldo na tabela MvPerCli
        lErro = CF("MvPerCli_Le_SldAnterior",dSaldo, lCliente, dtData, iFilialEmpresa, lComando)
        If lErro <> SUCESSO Then gError 76442

        'Se nao for o ultimo dia do mes, consulta a tabela MvDiaCli
        If Dias_Mes(iMes, iAno) <> Day(dtData) Then
        
            lErro = CF("MvDiaCli_Le_SldAnterior",dSaldoDiario, lCliente, dtData, iFilialEmpresa, lComando2)
            If lErro <> SUCESSO Then gError 76446
            
            dSaldo = dSaldo + dSaldoDiario
        
        End If
        
        iFilialEmpresaAnt = iFilialEmpresa
        lClienteAnt = lCliente
        dtDataAnt = dtData
        dSaldoAnt = dSaldo
        
    End If
    
    dSaldo = dSaldoAnt
        
    Exit Function
    
    Cliente_Obtem_SaldoAnterior = SUCESSO
    
Erro_Cliente_Obtem_SaldoAnterior:

    Cliente_Obtem_SaldoAnterior = gErr
    
    Select Case gErr
        
        Case 76442, 76446
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153278)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComando2 <> 0 Then Call Comando_Fechar(lComando2)
    
End Sub
