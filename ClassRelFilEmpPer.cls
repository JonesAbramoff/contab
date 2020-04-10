VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFilEmpPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private iUltFilialEmpresa As Integer
Private dtUltData As Date

Private dUltQuantidadeFaturada As Double
Private dUltValorFaturado As Double

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Public Function ObterQtdeFaturadaFilEmpPer(iFilialEmpresa As Integer, dtData As Date, dQuantidadeFaturada As Double) As Long
'Para obter a quantidade faturada de uma filialempresa em um periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaFilEmpPer
    
    If iFilialEmpresa <> iUltFilialEmpresa Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatFilEmp(iFilialEmpresa, dtData)
        If lErro <> SUCESSO Then gError 90197
        
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFaturadaFilEmpPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaFilEmpPer:

    ObterQtdeFaturadaFilEmpPer = gErr
    
    Select Case gErr

        Case 90197
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153122)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoFilEmpPer(iFilialEmpresa As Integer, dtData As Date, dValorFaturado As Double) As Long
'Para obter o valor faturado de uma filialempresa em um periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoFilEmpPer
    
    If iFilialEmpresa <> iUltFilialEmpresa Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatFilEmp(iFilialEmpresa, dtData)
        If lErro <> SUCESSO Then gError 90196
        
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFaturadoFilEmpPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorFaturadoFilEmpPer:

    ObterValorFaturadoFilEmpPer = gErr
    
    Select Case gErr

        Case 90196
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153123)

    End Select

    Exit Function

End Function

Private Function ObterRelMesFatFilEmp(iFilialEmpresa As Integer, dtData As Date)
'Para obter as quantidades e os valores a partir de slddiafat e torna-los disponiveis para as outras funcoes

Dim lErro As Long

Dim dtDataIni As Date
Dim dtDataFim As Date
Dim dQuantidadeFaturada As Double
Dim dValorFaturado As Double

On Error GoTo Erro_ObterRelMesFatFilEmp

    dtDataIni = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    dtDataFim = DateAdd("m", 1, dtDataIni) - 1

    lErro = Obter_QtdeValorFat_FilEmp(iFilialEmpresa, dtDataIni, dtDataFim, dQuantidadeFaturada, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90240
    
    iUltFilialEmpresa = iFilialEmpresa
    dtUltData = dtData
        
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltValorFaturado = dValorFaturado

    ObterRelMesFatFilEmp = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesFatFilEmp:

    ObterRelMesFatFilEmp = gErr
    
    Select Case gErr

        Case 90240
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153124)

    End Select

    Exit Function
    
End Function

Function Obter_QtdeValorFat_FilEmp(iFilialEmpresa As Integer, dtDataIni As Date, dtDataFim As Date, dQuantidadeFaturada As Double, dValorFaturado As Double, lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_QtdeValorFat_FilEmp

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90193
    End If
    
    'Acumula a Quantidade e o Valor Faturado da tabela SldDiaFat dentro do mes da data passada por parametro.
    lErro = Comando_Executar(lComando, "SELECT SUM(QuantFaturada), SUM(ValorFaturado) FROM SldDiaFat WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ?", dQuantidadeFaturada, dValorFaturado, iFilialEmpresa, dtDataIni, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 90194
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 90195
    
    Obter_QtdeValorFat_FilEmp = SUCESSO
    
    Exit Function
    
Erro_Obter_QtdeValorFat_FilEmp:

    Obter_QtdeValorFat_FilEmp = gErr
    
    Select Case gErr

        Case 90193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 90194, 90195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153125)

    End Select

    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function
