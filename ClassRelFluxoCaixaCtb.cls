VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFluxoCaixaCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bValoresObtidos As Boolean
Private adValor(1 To 6) As Double

Function ObterInfo(dValor As Double, ByVal iTipoInfo As Integer, ByVal lNumIntRel As Long) As Long

Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_ObterInfo

    If bValoresObtidos = False Then
    
        'Abre comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 184260
        Next
        
        'obter multa, juros e desconto das baixaspag
        lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorMulta), SUM(ValorJuros), SUM(ValorDesconto) FROM relfluxocaixactbbxpagaux WHERE NumIntRel = ?", _
            adValor(1), adValor(2), adValor(3), lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 184261
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 184262
        
        'obter multa, juros e desconto das baixasrec
        lErro = Comando_Executar(alComando(2), "SELECT SUM(ValorMulta), SUM(ValorJuros), SUM(ValorDesconto) FROM relfluxocaixactbbxrecaux WHERE NumIntRel = ?", _
            adValor(4), adValor(5), adValor(6), lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 184263
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 184264
        
        'Fecha comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    dValor = adValor(iTipoInfo)
    
    ObterInfo = SUCESSO
    
    Exit Function
    
Erro_ObterInfo:

    ObterInfo = gErr
    
    Select Case gErr
    
        Case 184260
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184261 To 184264
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELFLUXOCTBINFO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184265)

    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Sub Class_Initialize()
    bValoresObtidos = False
End Sub
