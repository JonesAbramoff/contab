VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRendAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Identificação
Private iUltFilialEmpresa As Integer
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date

'Valores armazenados na pesquisa anterior
Private dUltRendimentos As Double

Public Function ObterRendAplicPer(dRendimentos As Double, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor dos rendimentos das aplicacoes de uma filialempresa num periodo

Dim lErro As Long, dRendimentosAux As Double

On Error GoTo Erro_ObterRendAplicPer

    If dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then

        lErro = CF("Aplicacoes_ObterRendPer", iFilialEmpresa, dtDataInicial, dtDataFinal, dRendimentosAux)
        If lErro <> SUCESSO Then gError 184222
        
        dtUltDataInicial = dtDataInicial
        dtUltDataFinal = dtDataFinal
        iUltFilialEmpresa = iFilialEmpresa
        
        dUltRendimentos = dRendimentosAux
        
    End If
    
    dRendimentos = dUltRendimentos
    
    ObterRendAplicPer = SUCESSO
    
    Exit Function
    
Erro_ObterRendAplicPer:

     ObterRendAplicPer = gErr

    Select Case gErr

        Case 184222
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184223)

    End Select
    
    Exit Function
    
End Function

