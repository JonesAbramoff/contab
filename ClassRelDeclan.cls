VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDeclan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lComandoInicial As Long
Private lComandoFinal As Long
Private iFilialEmpresaAnt As Integer
Private dtDataInicialAnt As Date
Private dtDataFinalAnt As Date
Private dValorInicialAnt As Double
Private dValorFinalAnt As Double

Function Declan_EstoqueInicial(dValor As Double, dtDataInicial As Date, iFilialEmpresa As Integer) As Long
'Passa a data e a Filial Empresa e retorno o Valor do Estoque da Filial

Dim lErro As Long

On Error GoTo Erro_Declan_EstoqueInicial
        
    'Se não mudou os parametros
    If dtDataInicialAnt <> dtDataInicial Or iFilialEmpresaAnt <> iFilialEmpresa Then
                    
        'Lê o Valor do estoque na data passada para a FilialEmpresa
        lErro = CF("Declan_Obter_Estoque",dtDataInicial, iFilialEmpresa, dValor, lComandoInicial)
        If lErro <> SUCESSO Then gError 78074
        
        dValorInicialAnt = dValor
        dtDataInicialAnt = dtDataInicial
        iFilialEmpresaAnt = iFilialEmpresa
        
    End If
    
    dValor = dValorInicialAnt
        
    Exit Function
    
    Declan_EstoqueInicial = SUCESSO
    
Erro_Declan_EstoqueInicial:

    Declan_EstoqueInicial = gErr
    
    Select Case gErr
        
        Case 78074
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153082)

    End Select

    Exit Function

End Function

Function Declan_EstoqueFinal(dValor As Double, dtDataFinal As Date, iFilialEmpresa As Integer) As Long
'Passa a data e a Filial Empresa e retorno o Valor do Estoque da Filial

Dim lErro As Long

On Error GoTo Erro_Declan_EstoqueFinal
        
    'Se não mudou os parametros
    If dtDataFinalAnt <> dtDataFinal Or iFilialEmpresaAnt <> iFilialEmpresa Then
                    
        'Lê o Valor do estoque na data passada para a FilialEmpresa
        lErro = CF("Declan_Obter_Estoque",dtDataFinal, iFilialEmpresa, dValor, lComandoFinal)
        If lErro <> SUCESSO Then gError 78075
        
        dValorFinalAnt = dValor
        dtDataFinalAnt = dtDataFinal
        iFilialEmpresaAnt = iFilialEmpresa
        
    End If
    
    dValor = dValorFinalAnt
        
    Exit Function
    
    Declan_EstoqueFinal = SUCESSO
    
Erro_Declan_EstoqueFinal:

    Declan_EstoqueFinal = gErr
    
    Select Case gErr
        
        Case 78075
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153083)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComandoInicial <> 0 Then Call Comando_Fechar(lComandoInicial)
    If lComandoFinal <> 0 Then Call Comando_Fechar(lComandoFinal)
    
End Sub
