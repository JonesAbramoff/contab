VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoComissCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CROMATON_LIM_DIAS_CUSTOFIN = 14

Function Calcula_Mnemonico_ComissoesCust(ByVal objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux, ByVal objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, lRetorno As Long, dTaxaFin As Double

On Error GoTo Erro_Calcula_Mnemonico_ComissoesCust

    lRetorno = SUCESSO
    
    'Faz uma selecao pelo mnemonico
    Select Case objMnemonicoValor.sMnemonico
    
        Case "CustoFinanceiro"
            dTaxaFin = StrParaDbl(objMnemonicoValor.vParam(1))
            lErro = Comissoes_ObterCustoFinanceiro(objMnemonicoComissCalcAux, objMnemonicoValor, dTaxaFin)
            If lErro <> SUCESSO Then gError 130003
            
        Case Else
        
            lRetorno = MNEMONICOCOMISSOES_NAO_ENCONTRADO
            
    End Select
    
    Calcula_Mnemonico_ComissoesCust = lRetorno
     
    Exit Function
    
Erro_Calcula_Mnemonico_ComissoesCust:

    Calcula_Mnemonico_ComissoesCust = gErr
     
    Select Case gErr
          
        Case 130003
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function Comissoes_ObterCustoFinanceiro(ByVal objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux, ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal dTaxaFin As Double) As Long

Dim lErro As Long, objCondicaoPagto As New ClassCondicaoPagto, dDias As Double
Dim dTaxaFinDia As Double, dCustoFinanceiro As Double, iLinha As Integer

On Error GoTo Erro_Comissoes_ObterCustoFinanceiro

    'obtem os dados da condpagto
    lErro = objMnemonicoComissCalcAux.objTela.objMnemonicoComissCalc.CondPagto_Obtem_Dados(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 130004
    
    If objCondicaoPagto.iCodigo <> 0 Then
        
        lErro = CF("CalculaDias_CondPagamento", objCondicaoPagto, dDias)
        If lErro <> SUCESSO Then gError 130005
        
        If dDias > CROMATON_LIM_DIAS_CUSTOFIN Then
            Call objMnemonicoComissCalcAux.objTela.objMnemonicoComissCalc.Calcula_TaxaFinanceira(dDias, dTaxaFin, dTaxaFinDia)
        Else
            dTaxaFinDia = 0
        End If
            
    End If
    
    'Para cada linha no grid de itens
    For iLinha = 1 To objMnemonicoComissCalcAux.objTela.objGridItens.iLinhasExistentes
    
        dCustoFinanceiro = 1 / (1 + dTaxaFinDia)
        
        objMnemonicoValor.colValor.Add dCustoFinanceiro
    
    Next
    
    Comissoes_ObterCustoFinanceiro = SUCESSO
     
    Exit Function
    
Erro_Comissoes_ObterCustoFinanceiro:

    Comissoes_ObterCustoFinanceiro = gErr
     
    Select Case gErr
          
        Case 130004, 130005
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function
    
End Function
