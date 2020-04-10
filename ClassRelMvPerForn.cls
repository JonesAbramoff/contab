VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvPerForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim lComando As Long
Dim lComando2 As Long
Dim iFilialEmpresaAnt As Integer
Dim lFornecedorAnt As Long
Dim dtDataAnt As Date
Dim dSaldoAnt As Double

Function Fornecedor_Obtem_SaldoAnterior(dSaldo As Double, lFornecedor As Long, dtData As Date, iFilialEmpresa As Integer) As Long
'Retorna em dSaldo o SaldoContabil a receber do Fornecedor, consultando as tabelas MvPerForn e MvDiaForn

Dim lErro As Long
Dim iMes As Integer
Dim iAno As Integer
Dim dSaldoDiario As Double

On Error GoTo Erro_Fornecedor_Obtem_SaldoAnterior

    If iFilialEmpresa <> iFilialEmpresaAnt Or dtData <> dtDataAnt Or lFornecedor <> lFornecedorAnt Then
            
        lErro = CF("MvPerForn_Le_SldAnterior",dSaldo, lFornecedor, dtData, iFilialEmpresa, lComando)
        If lErro <> SUCESSO Then gError 76452
        
        'Se nao for o ultimo dia do mes, consulta a tabela MvDiaForn
        If Dias_Mes(iMes, iAno) <> Day(dtData) Then
        
            lErro = CF("MvDiaForn_Le_SldAnterior",dSaldoDiario, lFornecedor, dtData, iFilialEmpresa, lComando2)
            If lErro <> SUCESSO Then gError 76453
        
            dSaldo = dSaldo + dSaldoDiario
            
        End If
        
        iFilialEmpresaAnt = iFilialEmpresa
        lFornecedorAnt = lFornecedor
        dtDataAnt = dtData
        dSaldoAnt = dSaldo
        
    End If
    
    dSaldo = dSaldoAnt
        
    Exit Function
    
    Fornecedor_Obtem_SaldoAnterior = SUCESSO
    
Erro_Fornecedor_Obtem_SaldoAnterior:

    Fornecedor_Obtem_SaldoAnterior = gErr
    
    Select Case gErr
        
        Case 76452, 76453
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153282)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComando2 <> 0 Then Call Comando_Fechar(lComando2)
    
End Sub
