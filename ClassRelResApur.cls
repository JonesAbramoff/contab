VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelResApur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para guardar info das guias de recolhimento
Private mvarcolGuias As New Collection
Private dtDataDeAntGuias As Date
Private dtDataAteAntGuias As Date
Private iFilialEmpresaAntGuias As Date

'Para acumular os totais
Private dSubTotalCredito As Double
Private dSubTotalDebito As Double
Private dSaldoCredorAnt As Double

'Para ser usado na leitura das saidas
Private lComandoSaidas As Long
Private dtDataDeAntSaidas As Date
Private dtDataAteAntSaidas As Date
Private iFilialEmpresaAntSaidas As Date
Private dValorSaidas As Double

'Para ser usado na leitura das Entradas
Private lComandoEntradas As Long
Private dtDataDeAntEntradas As Date
Private dtDataAteAntEntradas As Date
Private iFilialEmpresaAntEntradas As Date
Private dValorEntrada As Double

Private lComandoSaldoCredor As Long
Private dtDataDeAntSaldoCredor As Date
Private dtDataAteAntSaldoCredor As Date
Private iFilialEmpresaAntSaldoCredor As Date

Function SubTotalDebito(dSubTotalRecebido As Double, dSubTotalPassado As Double) As Long

    dSubTotalPassado = dSubTotalRecebido
    dSubTotalDebito = dSubTotalRecebido
    
End Function

Function SubTotalCredito(dSubTotalRecebido As Double, dSubTotalPassado As Double) As Long

    dSubTotalPassado = dSubTotalRecebido
    dSubTotalCredito = dSubTotalRecebido
    
End Function

Function ApuracaoRes_SaldoDev(dSubTotalPassado As Double) As Long

    If (dSubTotalDebito - (dSubTotalCredito + dSaldoCredorAnt)) > 0 Then
        dSubTotalPassado = dSubTotalDebito - (dSubTotalCredito + dSaldoCredorAnt)
    Else
        dSubTotalPassado = 0
    End If
    
End Function

Function ApuracaoICMSRes_ImpRec(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

    
End Function

Function ApuracaoRes_ImpRecolhe(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

    
End Function

Function ApuracaoRes_SaldoCred(dSubTotalPassado As Double) As Long

    If ((dSubTotalCredito + dSaldoCredorAnt) - dSubTotalDebito) > 0 Then
        dSubTotalPassado = (dSubTotalCredito + dSaldoCredorAnt) - dSubTotalDebito
    Else
        dSubTotalPassado = 0
    End If
    
End Function

Function ApuracaoICMSRes_Saidas(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double

On Error GoTo Erro_ApuracaoICMSRes_Saidas
    
    If dtDataDeAntSaidas <> dtDataDe Or dtDataAteAntSaidas <> dtDataAte Or iFilialEmpresaAntSaidas <> iFilialEmpresa Then
            
        lErro = CF("LivRegES_Le_Saidas", dValor, dValorAux, dtDataDe, dtDataAte, iFilialEmpresa, lComandoSaidas)
        If lErro <> SUCESSO Then gError 78035
            
        dValorSaidas = dValor
        dtDataDeAntSaidas = dtDataDe
        dtDataAteAntSaidas = dtDataAte
        iFilialEmpresaAntSaidas = iFilialEmpresa
        
    End If
    
    dValor = dValorSaidas
    
    Exit Function
    
    ApuracaoICMSRes_Saidas = SUCESSO
    
Erro_ApuracaoICMSRes_Saidas:

    ApuracaoICMSRes_Saidas = gErr
    
    Select Case gErr
        
        Case 78035
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153338)

    End Select

    Exit Function

End Function

Function ApuracaoICMSRes_Entradas(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double

On Error GoTo Erro_ApuracaoICMSRes_Entradas
    
    If dtDataDeAntEntradas <> dtDataDe Or dtDataAteAntEntradas <> dtDataAte Or iFilialEmpresaAntEntradas <> iFilialEmpresa Then
            
        lErro = CF("LivRegES_Le_Entradas", dValor, dValorAux, dtDataDe, dtDataAte, iFilialEmpresa, lComandoEntradas)
        If lErro <> SUCESSO Then gError 78040
        
        dValorEntrada = dValor
        dtDataDeAntEntradas = dtDataDe
        dtDataAteAntEntradas = dtDataAte
        iFilialEmpresaAntEntradas = iFilialEmpresa
        
    End If
    
    dValor = dValorEntrada
        
    Exit Function
    
    ApuracaoICMSRes_Entradas = SUCESSO
    
Erro_ApuracaoICMSRes_Entradas:

    ApuracaoICMSRes_Entradas = gErr
    
    Select Case gErr
        
        Case 78040
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153339)

    End Select

    Exit Function

End Function

Function ApuracaoRes_Saidas(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double

On Error GoTo Erro_ApuracaoRes_Saidas
    
    If dtDataDeAntSaidas <> dtDataDe Or dtDataAteAntSaidas <> dtDataAte Or iFilialEmpresaAntSaidas <> iFilialEmpresa Then
            
        lErro = CF("LivRegES_Le_Saidas", dValorAux, dValor, dtDataDe, dtDataAte, iFilialEmpresa, lComandoSaidas)
        If lErro <> SUCESSO Then gError 78061
            
        dValorSaidas = dValor
        dtDataDeAntSaidas = dtDataDe
        dtDataAteAntSaidas = dtDataAte
        iFilialEmpresaAntSaidas = iFilialEmpresa
        
    End If
    
    dValor = dValorSaidas
    
    Exit Function
    
    ApuracaoRes_Saidas = SUCESSO
    
Erro_ApuracaoRes_Saidas:

    ApuracaoRes_Saidas = gErr
    
    Select Case gErr
        
        Case 78061
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153340)

    End Select

    Exit Function

End Function

Function ApuracaoRes_Entradas(dValor As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double

On Error GoTo Erro_ApuracaoRes_Entradas
    
    If dtDataDeAntEntradas <> dtDataDe Or dtDataAteAntEntradas <> dtDataAte Or iFilialEmpresaAntEntradas <> iFilialEmpresa Then
            
        lErro = CF("LivRegES_Le_Entradas", dValorAux, dValor, dtDataDe, dtDataAte, iFilialEmpresa, lComandoEntradas)
        If lErro <> SUCESSO Then gError 78062
        
        dValorEntrada = dValor
        dtDataDeAntEntradas = dtDataDe
        dtDataAteAntEntradas = dtDataAte
        iFilialEmpresaAntEntradas = iFilialEmpresa
        
    End If
    
    dValor = dValorEntrada
        
    Exit Function
    
    ApuracaoRes_Entradas = SUCESSO
    
Erro_ApuracaoRes_Entradas:

    ApuracaoRes_Entradas = gErr
    
    Select Case gErr
        
        Case 78062
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153341)

    End Select

    Exit Function

End Function

Function ApuracaoICMSRes_SalCredAnt(dSaldo As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_ApuracaoICMSRes_SalCredAnt
    
    If dtDataDeAntSaldoCredor <> dtDataDe Or dtDataAteAntSaldoCredor <> dtDataAte Or iFilialEmpresaAntSaldoCredor <> iFilialEmpresa Then
        
        objApuracao.iFilialEmpresa = iFilialEmpresa
        
        'Alteracao para ler o valor correto Tulio151002
        objApuracao.dtDataInicial = dtDataDe
        objApuracao.dtDataFinal = dtDataAte
        
        'Le o Saldo Credor do último Período
        lErro = CF("ApuracaoICMS_Le_SaldoCredor", objApuracao, lComandoSaldoCredor)
        If lErro <> SUCESSO And lErro <> 78049 Then gError 78045
        
        dSaldoCredorAnt = objApuracao.dSaldoCredorInicial
        dtDataDeAntSaldoCredor = dtDataDe
        dtDataAteAntSaldoCredor = dtDataAte
        iFilialEmpresaAntSaldoCredor = iFilialEmpresa
        
        
    End If
    
    dSaldo = dSaldoCredorAnt
        
    Exit Function
    
    ApuracaoICMSRes_SalCredAnt = SUCESSO
    
Erro_ApuracaoICMSRes_SalCredAnt:

    ApuracaoICMSRes_SalCredAnt = gErr
    
    Select Case gErr
        
        Case 78045
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153342)

    End Select

    Exit Function

End Function

Function ApuracaoRes_SalCredAnt(dSaldo As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim dValorAux As Double
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_ApuracaoRes_SalCredAnt
    
    If dtDataDeAntSaldoCredor <> dtDataDe Or dtDataAteAntSaldoCredor <> dtDataAte Or iFilialEmpresaAntSaldoCredor <> iFilialEmpresa Then
        
        objApuracao.iFilialEmpresa = iFilialEmpresa
        
        'Alteracao para ler o valor correto Tulio151002
        objApuracao.dtDataInicial = dtDataDe
        objApuracao.dtDataFinal = dtDataAte
        
        'Le o Saldo Credor do último Período
        lErro = CF("ApuracaoIPI_Le_SaldoCredor", objApuracao, lComandoSaldoCredor)
        If lErro <> SUCESSO And lErro <> 78049 Then gError 88022
        
        dSaldoCredorAnt = objApuracao.dSaldoCredorInicial
        dtDataDeAntSaldoCredor = dtDataDe
        dtDataAteAntSaldoCredor = dtDataAte
        iFilialEmpresaAntSaldoCredor = iFilialEmpresa
        
        
    End If
    
    dSaldo = dSaldoCredorAnt
        
    Exit Function
    
    ApuracaoRes_SalCredAnt = SUCESSO
    
Erro_ApuracaoRes_SalCredAnt:

    ApuracaoRes_SalCredAnt = gErr
    
    Select Case gErr
        
        Case 88022
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153343)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComandoEntradas <> 0 Then Call Comando_Fechar(lComandoEntradas)
    If lComandoSaidas <> 0 Then Call Comando_Fechar(lComandoSaidas)
    If lComandoSaldoCredor <> 0 Then Call Comando_Fechar(lComandoSaldoCredor)
    
    Set mvarcolGuias = Nothing
    
End Sub

Function GuiaNumero(sNumero As String, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaNumero

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124100
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        sNumero = ""
    Else
        sNumero = mvarcolGuias(iIndice).sNumero
    End If
    
    GuiaNumero = SUCESSO
     
    Exit Function
    
Erro_GuiaNumero:

    GuiaNumero = gErr
     
    Select Case gErr
          
        Case 124100
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153344)
     
    End Select
     
    Exit Function

End Function

Function GuiaData(dtData As Date, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaData

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124101
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        dtData = DATA_NULA
    Else
        dtData = mvarcolGuias(iIndice).dtData
    End If
    
    GuiaData = SUCESSO
     
    Exit Function
    
Erro_GuiaData:

    GuiaData = gErr
     
    Select Case gErr
          
        Case 124101
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153345)
     
    End Select
     
    Exit Function

End Function

Function GuiaValor(dValor As Double, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaValor

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124102
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        dValor = 0
    Else
        dValor = mvarcolGuias(iIndice).dValor
    End If
    
    GuiaValor = SUCESSO
     
    Exit Function
    
Erro_GuiaValor:

    GuiaValor = gErr
     
    Select Case gErr
          
        Case 124102
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153346)
     
    End Select
     
    Exit Function

End Function

Function GuiaOrgao(sOrgao As String, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaOrgao

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124103
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        sOrgao = ""
    Else
        sOrgao = mvarcolGuias(iIndice).sOrgaoArrecadador
    End If
    
    GuiaOrgao = SUCESSO
     
    Exit Function
    
Erro_GuiaOrgao:

    GuiaOrgao = gErr
     
    Select Case gErr
          
        Case 124103
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153347)
     
    End Select
     
    Exit Function

End Function

Function GuiaDataEnt(dtDataEnt As Date, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaDataEnt

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124104
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        dtDataEnt = DATA_NULA
    Else
        dtDataEnt = mvarcolGuias(iIndice).dtDataEntrega
    End If
        
    GuiaDataEnt = SUCESSO
     
    Exit Function
    
Erro_GuiaDataEnt:

    GuiaDataEnt = gErr
     
    Select Case gErr
          
        Case 124104
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153348)
     
    End Select
     
    Exit Function

End Function

Function GuiaLocal(sLocal As String, ByVal iIndice As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_GuiaLocal

    If dtDataDeAntGuias <> dtDataDe Or dtDataAteAntGuias <> dtDataAte Or iFilialEmpresaAntGuias <> iFilialEmpresa Then
    
        lErro = CF("ApuracaoRes_ObtemGuias", dtDataDe, dtDataAte, iFilialEmpresa, mvarcolGuias)
        If lErro <> SUCESSO Then gError 124105
        
        dtDataDeAntGuias = dtDataDe
        dtDataAteAntGuias = dtDataAte
        iFilialEmpresaAntGuias = iFilialEmpresa
    
    End If
    
    If iIndice < 1 Or iIndice > mvarcolGuias.Count Then
        sLocal = ""
    Else
        sLocal = mvarcolGuias(iIndice).sLocalEntrega
    End If
    
    GuiaLocal = SUCESSO
     
    Exit Function
    
Erro_GuiaLocal:

    GuiaLocal = gErr
     
    Select Case gErr
          
        Case 124105
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153349)
     
    End Select
     
    Exit Function

End Function
